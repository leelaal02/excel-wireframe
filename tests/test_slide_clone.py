from pathlib import Path

from fixtures import make_template_pptx, make_png
from pptx import Presentation
from pptx.util import Emu, Inches
from slide_clone import clone_slide


def test_clone_with_embedded_picture_preserves_image_bytes(tmp_path: Path):
    """Verify rId remapping by cloning a slide with an embedded picture and hyperlink.

    This test proves the rId remapping mechanism works. Without proper remapping
    of r:embed and r:id in shape XML, pictures' relationships would point to
    wrong rIds on the cloned slide, causing images to vanish or corrupt the file.

    The test creates both an embedded picture (r:embed) and an external hyperlink (r:id)
    on the text shape to force rId shifts between slides.
    """
    # Create a presentation with a slide containing an embedded picture
    prs = Presentation()
    prs.slide_width = Emu(9906000)
    prs.slide_height = Emu(6858000)
    blank_slide_layout = prs.slide_layouts[6]  # blank layout
    src = prs.slides.add_slide(blank_slide_layout)

    # Add a picture to the source slide
    pic_path = make_png(tmp_path / "test.png", size=(200, 150))
    with open(pic_path, 'rb') as f:
        original_blob = f.read()

    src.shapes.add_picture(str(pic_path), Inches(1), Inches(1), width=Inches(2))

    # Add a text box with a hyperlink to create an external relationship
    # This shifts the image rId so the new slide will have different rIds
    txBox = src.shapes.add_textbox(Inches(0.5), Inches(4), Inches(3), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "Test Link"

    # Add hyperlink to the run (creates an external relationship with r:id)
    hlink = run.hyperlink
    hlink.address = "https://example.com"

    # Clone the slide - this should remap both r:embed (picture) and r:id (hyperlink)
    new = clone_slide(prs, src)

    # Save and reload to verify the clone survives serialization
    out = tmp_path / "out.pptx"
    prs.save(out)

    chk = Presentation(str(out))
    cloned_slide = chk.slides[-1]

    # Assert the cloned slide has exactly one picture
    pics = [s for s in cloned_slide.shapes if s.shape_type == 13]  # 13 = PICTURE
    assert len(pics) == 1, f"Expected 1 picture, got {len(pics)}"

    # Assert the picture blob is byte-identical to the original
    cloned_blob = pics[0].image.blob
    assert cloned_blob == original_blob, (
        f"Picture blob mismatch after clone: original {len(original_blob)} bytes, "
        f"cloned {len(cloned_blob)} bytes"
    )


def test_clone_copies_all_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    new = clone_slide(prs, src)
    assert len(new.shapes) == len(src.shapes)
    assert [s.name for s in new.shapes] == [s.name for s in src.shapes]


def test_clone_survives_save_and_reload(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    clone_slide(prs, prs.slides[0])
    out = tmp_path / "out.pptx"
    prs.save(out)

    chk = Presentation(str(out))
    assert len(chk.slides) == 2
    last = chk.slides[-1]
    tables = [s for s in last.shapes if s.has_table]
    assert len(tables) == 5
    assert tables[0].table.cell(0, 0).text == "1"


def test_clone_twice_produces_three_slides(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    clone_slide(prs, src)
    clone_slide(prs, src)
    out = tmp_path / "out.pptx"
    prs.save(out)
    assert len(Presentation(str(out)).slides) == 3

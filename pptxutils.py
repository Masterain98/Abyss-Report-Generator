from pptx.shapes.base import BaseShape


def shape_alt_text(shape: BaseShape) -> str:
    # https://github.com/scanny/python-pptx/pull/512#issuecomment-1713100069
    """Alt-text defined in shape's `desc` attribute, or "" if not present."""
    return shape._element._nvXxPr.cNvPr.attrib.get("descr", "")

"""Convert images into Excel pixel art."""

__all__ = ["image_to_excel"]


def __getattr__(name: str):
    if name == "image_to_excel":
        from .converter import image_to_excel

        return image_to_excel
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

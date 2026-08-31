# E's VRS Database Updater - Python Port
__version__ = "2.1.0"


def display_version() -> str:
    """Short version for window titles and banners: "2.1.0" -> "v2.1".

    Everything user-facing derives from __version__ so the number cannot be
    bumped in one place and left stale in another.
    """
    return "v" + ".".join(__version__.split(".")[:2])

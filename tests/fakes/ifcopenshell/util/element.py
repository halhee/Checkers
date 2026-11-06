def get_psets(element, **kwargs):
    return getattr(element, "_psets", {}) or {}

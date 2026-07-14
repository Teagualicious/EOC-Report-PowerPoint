"""Resolve which shape a mapping entry targets.

Mappings historically identified shapes by positional index alone, so any
edit to a template (shape added, deleted, or reordered) silently pointed
existing assignments at the wrong shapes. Mapping entries now also store
the persistent PowerPoint shape id (``shape_uid``, python-pptx
``shape.shape_id`` / COM ``Shape.Id``) and the shape name; this module
resolves them back to a position in a live shape collection.

Resolution order: uid match, then unique-name match, then — only for
legacy entries that carry no stored identity — the positional index. An
entry whose stored identity matches nothing returns None: skipping (and
reporting) beats writing into whatever shape now occupies the old index.

Pure Python (no pptx/COM imports) so the python-pptx fill engine and the
Windows COM live preview share one implementation and CI exercises it.
"""


def resolve_shape_index(id_name_pairs, mapped_index, shape_uid=None,
                        shape_name=None):
    """Resolve a mapping entry to an index into a shape collection.

    id_name_pairs: [(uid, name), ...] in collection order (0-based).
    Returns the resolved 0-based index, or None when the target shape
    cannot be safely identified.
    """
    if shape_uid is not None:
        for i, (uid, _name) in enumerate(id_name_pairs):
            if uid == shape_uid:
                return i
    if shape_name:
        hits = [i for i, (_uid, name) in enumerate(id_name_pairs)
                if name == shape_name]
        if len(hits) == 1:  # duplicate names never match — no guessing
            return hits[0]
    if shape_uid is None and not shape_name:
        # Legacy entry (pre-uid mapping): positional identity is all there is.
        if 0 <= mapped_index < len(id_name_pairs):
            return mapped_index
    return None

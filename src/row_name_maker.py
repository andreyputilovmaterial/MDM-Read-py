"""Pure row-name builders for the MDD reader.

The reader flattens the nested MDM object model into flat rows, each with a unique
``name``. The exact spelling of that ``name`` is a *row name convention*:

- ``ap`` (default): the parent's local name + ``.`` + the child's local name.
  Categories are written as ``Parent.Categories[Cat]``. When ``stepover`` prints a
  shared-list reference instead of expanding it, it is written as
  ``Parent.Categories[@SL_Name]``.
- ``mdm`` (placeholder): the child's ``FullName`` is used as-is, assuming it already
  contains the full dotted path (no parent prefix is appended).

This module is pure string/attribute manipulation: it only reads ``.Name`` /
``.FullName`` / ``.IsReference`` off the objects it is handed, so it can be unit-tested
without the COM runtime. It must never import ``win32com`` or any sibling module.
"""


# Single source of truth for the available conventions. ``ROW_NAME_CONVENTIONS.keys()``
# is used to populate the CLI ``--config-row-name-convention`` choices, so adding a convention
# here is enough to expose it (no extra hardcoded strings elsewhere).
ROW_NAME_CONVENTIONS = {
    'ap': {
        'mode': 'compose',        # build the name from parent + child parts
        'separator': '.',
        'category_token': 'Categories',
        'open': '[',
        'close': ']',
    },
    'mdm': {
        'mode': 'fullname',       # use the item's FullName as-is (placeholder)
        'separator': '.',         # informational only - FullName is expected to be dotted
    },
}

DEFAULT_ROW_NAME_CONVENTION = 'ap'

# Prefix used when ``stepover`` prints a shared-list reference instead of expanding it.
# Deliberately not exposed as a CLI flag; it can still be overridden through the reader
# config (``config_row_name_reference_prefix``).
DEFAULT_REFERENCE_PREFIX = '@'


def resolve_convention(convention):
    """Return the convention definition for ``convention`` or fail early.

    Accepts a convention name (string) or a full convention dict (for ad-hoc use).
    """
    if convention is None:
        convention = DEFAULT_ROW_NAME_CONVENTION
    if isinstance(convention, str):
        if convention not in ROW_NAME_CONVENTIONS:
            raise ValueError(
                f'unrecognized row name convention: "{convention}". '
                f'Available: {", ".join(sorted(ROW_NAME_CONVENTIONS))}'
            )
        return ROW_NAME_CONVENTIONS[convention]
    return convention


def _is_fullname_mode(convention):
    return resolve_convention(convention).get('mode') == 'fullname'


def _local_name(item):
    """Return an item's local ``Name`` (falling back to its string form)."""
    return f'{item.Name}' if hasattr(item, 'Name') else f'{item}'


def _full_name(item):
    """Return an item's ``FullName`` (falling back to its local ``Name``).

    TODO (placeholder for the ``mdm`` convention): review whether falling back to
    ``Name`` is the right behavior for items that do not expose ``FullName``.
    """
    return f'{item.FullName}' if hasattr(item, 'FullName') else _local_name(item)


def is_reference(item):
    """Return whether an item is a shared-list reference (``stepover``)."""
    try:
        return bool(item.IsReference)
    except AttributeError:
        return False


def make_root_name(item, convention=None):
    """Name a top-level row (the item itself)."""
    if _is_fullname_mode(convention):
        return _full_name(item)
    return _local_name(item)


def make_child_name(parent, child, convention=None):
    """Name a nested row (subfield / page entry / helper field).

    ``parent`` may be an item object or an already-composed parent name string.
    """
    if _is_fullname_mode(convention):
        return _full_name(child)
    convention = resolve_convention(convention)
    return f'{_local_name(parent)}{convention["separator"]}{_local_name(child)}'


def make_category_name(parent, child, child_name=None, is_reference=False,
                       reference_prefix=None, convention=None):
    """Name a category/element row.

    ``child_name`` is the resolved category name (usually ``item.Name``; in
    ``stepover`` mode it is the referenced shared list's name). When ``is_reference``
    is true (``stepover``) the name is prefixed with ``reference_prefix`` (``@`` by
    default).
    """
    if _is_fullname_mode(convention):
        # TODO (placeholder for the ``mdm`` convention): review category / shared-list
        # naming in fullname mode.
        return _full_name(child)
    convention = resolve_convention(convention)
    if child_name is None:
        child_name = _local_name(child)
    prefix = ''
    if is_reference:
        prefix = reference_prefix if reference_prefix is not None else DEFAULT_REFERENCE_PREFIX
    return (
        f'{_local_name(parent)}'
        f'{convention["separator"]}'
        f'{convention["category_token"]}'
        f'{convention["open"]}'
        f'{prefix}{child_name}'
        f'{convention["close"]}'
    )

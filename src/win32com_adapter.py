"""Lazy adapter around ``win32com.client`` so the MDD reader can import on any OS.

``win32com.client`` is only available on Windows with pywin32 installed. Importing it
lazily here lets the rest of the code import this module anywhere, and only fail with a
readable message when a COM object is actually requested.
"""


import sys



# STDOUT_COLOR_RED = "\033[91m"
STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"
STDOUT_COLOR_GREEN = "\033[32m"





class Win32ComUnavailable(RuntimeError):
    """Raised (historically) when COM/pywin32 is unavailable on this platform."""
    pass


class Win32Client:
    """Thin wrapper over ``win32com.client`` that imports it lazily and fails early.

    Use ``Win32Client.dispatch(prog_id)`` (or the legacy ``Win32Client.Dispatch``) to
    create a COM object.
    """
    _client = None
    _import_error = None

    @classmethod
    def _load(cls):
        """Import ``win32com.client`` once and cache it (or the import error)."""
        if cls._client is not None:
            return

        try:
            import win32com.client
            cls._client = win32com.client
        except Exception as e:
            cls._import_error = e

    @classmethod
    def dispatch(cls, prog_id):
        """Create and return the COM object for ``prog_id``.

        Raises a readable error (printed to stderr) if pywin32 is missing or COM is
        not supported on this platform.
        """
        cls._load()

        if cls._client is None:
            if sys.platform.startswith("win"):
                # raise Win32ComUnavailable(
                #     "Please install pywin32 module; win32.client is required on Windows but could not be imported"
                # ) from cls._import_error
                print(f'{STDOUT_COLOR_RED}Please install pywin32 module; win32.client is required on Windows but could not be imported{STDOUT_COLOR_RESET}',file=sys.stderr)
                raise cls._import_error
            else:
                print(f'{STDOUT_COLOR_RED}COM is not supported on this platform{STDOUT_COLOR_RESET}',file=sys.stderr)
                # raise Win32ComUnavailable(
                #     "COM is not supported on this platform"
                # ) from cls._import_error
                raise cls._import_error

        return cls._client.Dispatch(prog_id)

    @classmethod
    def Dispatch(self,*args):
        """Legacy alias for :meth:`dispatch`."""
        return self.dispatch(*args)

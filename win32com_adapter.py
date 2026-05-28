import sys



# STDOUT_COLOR_RED = "\033[91m"
STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"
STDOUT_COLOR_GREEN = "\033[32m"





class Win32ComUnavailable(RuntimeError):
    pass


class Win32Client:
    _client = None
    _import_error = None

    @classmethod
    def _load(cls):
        if cls._client is not None:
            return

        try:
            import win32com.client
            cls._client = win32com.client
        except Exception as e:
            cls._import_error = e

    @classmethod
    def dispatch(cls, prog_id):
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
        return self.dispatch(*args)

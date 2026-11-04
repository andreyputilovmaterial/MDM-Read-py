"""Universal dispatcher that routes ``--program`` to the right entry point.

Used both during development and from the pinliner bundle (``launcher.main()``).
"""



import argparse
# from pathlib import Path
import traceback, sys
import importlib






# Entry-point modules (``read_mdd``, the reporter, and the generated version string)
# are imported lazily inside their ``call_*`` handlers via :func:`_import_module`, so
# ``--program done`` works as a dependency-free smoke test in a fresh checkout.




# STDOUT_COLOR_RED = "\033[91m"
STDOUT_COLOR_RED = "\033[31m"
STDOUT_COLOR_RESET = "\033[0m"
STDOUT_COLOR_GREEN = "\033[32m"




def _import_module(flat_name):
    # Lazily import an entry-point module. Imports are selected on `__package__`
    # (set reliably by the import system in every case), not `__name__`:
    #
    #   * `__package__` truthy  ->  relative import (`python -m src.launcher`,
    #     or imported via the pinliner bundle).
    #   * `__package__` falsy   ->  flat import (inlined with no parent package).
    if __package__:
        # package context
        return importlib.import_module('.' + flat_name, package=__package__)
    else:
        # flat context: no parent package
        return importlib.import_module(flat_name)


def _import_version():
    # The generated version string only exists after a build step has written
    # ``GENERATED/_VERSION.py``. Fall back to a placeholder so ``--program version``
    # still works (and ``--program done`` never needs it) in a fresh checkout.
    try:
        version_module = _import_module('GENERATED._VERSION')
    except ImportError:
        return 'unknown'
    return version_module._VERSION


def call_read_mdd_program(*argcs,**kwargs):
    read_mdd = _import_module('read_mdd')
    return read_mdd.entry_point(*argcs,**kwargs)

def call_report_html_program(*argcs,**kwargs):
    report_html_create = _import_module('lib.mdmreportpy.src.report_create')
    return report_html_create.entry_point(*argcs,**kwargs)

def call_test_program(*argcs,**kwargs):
    msg = '''
hello, world! From mdd_read
    '''
    print(msg)
    return True

def call_done_program(*argcs,**kwargs):
    print(f'{STDOUT_COLOR_GREEN}done{STDOUT_COLOR_RESET}')
    return True

def call_printversion_program(*argcs,**kwargs):
    msg = _import_version().strip()
    print(msg)
    return True




run_programs = {
    'read_mdd': call_read_mdd_program,
    'report': call_report_html_program,
    'report_html': call_report_html_program,
    'test': call_test_program,
    'done': call_done_program,
    'version': call_printversion_program,
}



def main():
    """Parse ``--program`` and run the matching handler, formatting errors for users."""
    try:
        parser = argparse.ArgumentParser(
            description="Universal caller of mdmtoolsap-py utilities"
        )
        parser.add_argument(
            #'-1',
            '--program',
            choices=dict.keys(run_programs),
            type=str,
            required=True
        )
        args, args_rest = parser.parse_known_args()
        if args.program:
            program = '{arg}'.format(arg=args.program)
            if program in run_programs:
                run_programs[program](args_rest)
            else:
                raise AttributeError('program to run not recognized: {program}'.format(program=args.program))
        else:
            print('program to run not specified')
            raise AttributeError('program to run not specified')
    except Exception as e:
        # the program is designed to be user-friendly
        # that's why we reformat error messages a little bit
        # stack trace is still printed (I even made it longer to 20 steps!)
        # but the error message itself is separated and printed as the last message again

        # for example, I don't write "print('File Not Found!');exit(1);", I just write "raise FileNotFoundErro()"
        print('',file=sys.stderr)
        print('Stack trace:',file=sys.stderr)
        print('',file=sys.stderr)
        traceback.print_exception(e,limit=20)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print('',file=sys.stderr)
        print(f'{STDOUT_COLOR_RED}Error:{STDOUT_COLOR_RESET}',file=sys.stderr)
        print('',file=sys.stderr)
        print(f'{STDOUT_COLOR_RED}{e}{STDOUT_COLOR_RESET}',file=sys.stderr)
        print('',file=sys.stderr)
        exit(1)


if __name__ == '__main__':
    main()

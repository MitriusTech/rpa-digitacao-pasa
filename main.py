import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import sys
from packages.commons import *
from packages.bot_base import *
from wrapt_timeout_decorator import *
import features.core as core
from multiprocessing import Process, freeze_support

def handle_exceptions_with(excepthook, target, /, *args, **kwargs):
    try:
        target(*args, **kwargs)
    except:
        excepthook(*sys.exc_info())
        sys.exit(-1)

def main():
    core.run()
    return 0

if __name__ == '__main__':

    freeze_support()

    if len(sys.argv) > 1 and sys.argv[2] == "local":
        handle_exceptions_with(show_exception_and_exit, main) # debug mode
    else:
        find_and_terminate_other_instance()
        Process(target=handle_exceptions_with, args=(show_exception_and_exit, main)).start()
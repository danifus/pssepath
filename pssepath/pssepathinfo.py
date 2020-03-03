import logging

import pssepath
from pssepath.core import check_already_present_psse
from pssepath.compat import compat_input, simple_print


logger = logging.getLogger(__name__)


if __name__ == "__main__":
    # print the available psse installs.
    logging.basicConfig(format="%(message)s", level=logging.INFO)
    check_already_present_psse()
    simple_print("Found the following PSSE versions installed:\n")
    pssepath.print_psse_selection()
    simple_print("\n\nFound the following Python installations:")
    pssepath.print_python_selection()
    compat_input("Press Enter to continue...")

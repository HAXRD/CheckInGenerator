# Copyright (c) 2021, Xu Chen, FUNLab, Xiamen University
# All rights reserved.
# SPDX-License-Identifier: MIT
# For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/MIT

from datetime import datetime
import xlsxwriter
from pprint import pprint
from worker import Worker

if __name__ == "__main__":
    worker = Worker(2021, 3, 0.2)
    worker.write_checkin_xlsx()
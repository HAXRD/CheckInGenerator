# Copyright (c) 2021, Xu Chen, FUNLab, Xiamen University
# All rights reserved.
# SPDX-License-Identifier: MIT
# For full license text, see the LICENSE file in the repo root or https://opensource.org/licenses/MIT

from datetime import date, datetime, time, timedelta
import calendar
import random
import numpy as np
import xlsxwriter
import math

WEEKDAY_LOOKUP = [
    "一",
    "二",
    "三",
    "四",
    "五",
    "六",
    "日",
    "二"
]

VALID_RANGE = {
    "morning":   [time(8, 0, 1),  time(11,59,59)],
    "afternoon": [time(12, 0, 1), time(17,59,59)],
    "evening":   [time(18, 0, 1), time(22,29,59)]
}



class Worker:
    def __init__(self,
                 year, month, std_dev=0.1,
                 morning=True, afternoon=True, evening=True,
                 enable_occasional_checkin_n_late=True,
                 p_occasional_checkin_n_late=0.1,
                 weekend=False,
                 enable_occasional_weekend_checkin=True,
                 p_occasional_weekend_checkin=0.3
                 ):
        self.year, self.month, self.std_dev = year, month, std_dev
        self.morning, self.afternoon, self.evening = morning, afternoon, evening
        self.enable_occasional_checkin_n_late =enable_occasional_checkin_n_late
        self.p_occasional_checkin_n_late = p_occasional_checkin_n_late
        self.weekend = weekend
        self.enable_occasional_weekend_checkin = enable_occasional_weekend_checkin
        self.p_occasional_weekend_checkin = p_occasional_weekend_checkin

    def write_checkin_xlsx(self, output_dir="./", exception_days=[]):
        workbook = xlsxwriter.Workbook('tmp.xlsx')
        worksheet = workbook.add_worksheet()

        row, col = 1, 0
        data = self._populate_data(exception_days)
        time_format = workbook.add_format({'num_format': 'hh:mm'})
        for trans in data:
            worksheet.write_string(row, col, trans[0])
            for idx, item in enumerate(trans[1:]):
                if item == None:
                    worksheet.write_blank(row, col+idx+1, item)
                else:
                    print(item)
                    worksheet.write_datetime(row, col+idx+1, item, time_format)
            row += 1
        workbook.close()


    def _populate_data(self, exception_days=[]):
        data = []
        for day in range(1, calendar.monthrange(self.year, self.month)[1]+1):
            row = []
            # 1st column
            curr_date = date(self.year, self.month, day)
            curr_date_weekday = curr_date.weekday()
            date_n_weekday = curr_date.strftime("%d") + " " + str(WEEKDAY_LOOKUP[curr_date_weekday]) # 2021-04-01 "01 四"
            row.append(date_n_weekday)

            if curr_date_weekday <= 4 or random.random() < self.p_occasional_weekend_checkin:
                if day in exception_days:
                    for _ in range(6):
                        row.append(None)
                else:
                    # Populate morning
                    checkin, checkout = self._populate_a_session("morning", self.morning, self.p_occasional_checkin_n_late)
                    row.append(checkin)
                    row.append(checkout)
                    # Populate afternoon
                    checkin, checkout = self._populate_a_session("afternoon", self.afternoon, self.p_occasional_checkin_n_late)
                    row.append(checkin)
                    row.append(checkout)
                    # Populate evening
                    checkin, checkout = self._populate_a_session("evening", self.evening, self.p_occasional_checkin_n_late)
                    row.append(checkin)
                    row.append(checkout)
            data.append(row)
        return data

    def _populate_a_session(self, timeslot, enable, p):
        while True:
            checkin, checkout = None, None
            if enable or random.random() < p:
                checkin  = datetime.combine(date.today(), VALID_RANGE[timeslot][0]) + self._get_a_time_delta(timeslot)
                checkout = datetime.combine(date.today(), VALID_RANGE[timeslot][1]) - self._get_a_time_delta(timeslot)
            if not (checkin == None and checkout == None):
                return checkin, checkout

    def _get_a_time_delta(self, timeslot):
        while (True):
            multiplier = math.fabs(random.gauss(0, self.std_dev))
            delta_range_in_minutes = 60*(VALID_RANGE[timeslot][1].hour - VALID_RANGE[timeslot][0].hour) + \
                (VALID_RANGE[timeslot][1].minute - VALID_RANGE[timeslot][0].minute)
            if multiplier <= 1:
                time_delta = timedelta(
                    minutes=int(multiplier * delta_range_in_minutes))
                return time_delta
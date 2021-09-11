import openpyxl
import csv
import math
data_list = list()
org_name = ''
main_fire_data = {
    2011: [],
    2012: [],
    2013: [],
    2014: [],
    2015: [],
    2016: [],
    2017: [],
    2018: [],
    2019: [],
    2020: []
}
max_fire_data = {
    2011: [],
    2012: [],
    2013: [],
    2014: [],
    2015: [],
    2016: [],
    2017: [],
    2018: [],
    2019: [],
    2020: []
}
YEAR_LIST = [2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]


def get_true_format(item):
    if type(item) == str:
        if "," in item:
            return float(item.replace(",", "."))
        if item == '':
            return 0.00000000000000001
        return float(item)
    elif item is None:
        return 0
    return float(item)


class Fire:
    def __init__(self, fire_day_numder, year, ap, pa, rh, sm, tm, st):
        self.fire_day_numder = fire_day_numder
        self.year = year
        self.atmospheric_pressure = get_true_format(ap)
        self.precipitation_amount = get_true_format(pa)
        self.relative_humidity = get_true_format(rh)
        self.soil_moisture = get_true_format(sm)
        self.temperature = get_true_format(tm)
        self.soil_temperature = get_true_format(st)
        self.dew_point = self.get_dew_point()
        self.komplex_kof = self.get_dew_point()
        self.mod_kof = self.get_mod_kof()
        self.mod_pok = self.get_mod_pok(main_fire_data)

    def get_mod_kof(self):
        if self.precipitation_amount <= 0.5:
            return 1
        elif self.precipitation_amount <= 2:
            return 0.8
        elif self.precipitation_amount <= 5:
            return 0.4
        elif self.precipitation_amount <= 12:
            return 0.2
        elif self.precipitation_amount <= 19:
            return 0.1
        else:
            return 0

    def get_mod_pok(self, fire_data):
        if int(self.fire_day_numder) == 2:
            return (self.temperature - self.dew_point) * self.temperature
        else:
            last_day = fire_data[self.year][-1]
            return (last_day.mod_pok + (last_day.temperature - last_day.dew_point)
                    * last_day.temperature) * self.mod_kof

    def get_komplex_kof(self):
        if self.precipitation_amount <= 3:
            return 0
        else:
            return 1

    def get_dew_point(self):
        b = math.log10(self.relative_humidity / 4) + ((self.temperature / 4) * 7.45) / (235 + (self.temperature / 4)) - 2
        return (235 * b) / (7.45 - b)

    def get_full_fire_info(self):
        return self.fire_day_numder, self.atmospheric_pressure, self.precipitation_amount, self.relative_humidity,\
               self.soil_moisture, self.temperature, self.soil_temperature, self.dew_point, self.komplex_kof,\
               self.mod_kof, self.mod_pok


flag1 = True
flag2 = True
flag3 = True
with open('DATA_20210831_193317.csv', encoding="utf8") as f:
    for row in csv.reader(f, delimiter=';', quotechar='"'):
        if flag1:
            flag1 = False
            org_name = row[1]
            continue
        if flag2:
            flag2 = False
            continue
        if flag3:
            flag3 = False
            continue
        data_list.append(row)
for row_id in range(4, len(data_list), 4):
    main_row = data_list[row_id][0]
    row_2011_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][71], -111111111]
    row_2012_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][72], -111111111]
    row_2013_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][73], -111111111]
    row_2014_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][74], -111111111]
    row_2015_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][75], -111111111]
    row_2016_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][76], -111111111]
    row_2017_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][77], -111111111]
    row_2018_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][78], -111111111]
    row_2019_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][79], -111111111]
    row_2020_max = [-1111111111, -1111111111, -111111111, -11111111, data_list[row_id][80], -111111111]
    row_2011_main = [0, 0, 0, 0, data_list[row_id][71], 0]
    row_2012_main = [0, 0, 0, 0, data_list[row_id][72], 0]
    row_2013_main = [0, 0, 0, 0, data_list[row_id][73], 0]
    row_2014_main = [0, 0, 0, 0, data_list[row_id][74], 0]
    row_2015_main = [0, 0, 0, 0, data_list[row_id][75], 0]
    row_2016_main = [0, 0, 0, 0, data_list[row_id][76], 0]
    row_2017_main = [0, 0, 0, 0, data_list[row_id][77], 0]
    row_2018_main = [0, 0, 0, 0, data_list[row_id][78], 0]
    row_2019_main = [0, 0, 0, 0, data_list[row_id][79], 0]
    row_2020_main = [0, 0, 0, 0, data_list[row_id][80], 0]
    for test_row in data_list[row_id-1:row_id+2]:
        for i in range(1, 10):
            exec(f"row_201{i}_max[0] = max(row_201{i}_max[0], get_true_format(test_row[{i}]))")
        row_2020_max[0] = max(row_2020_max[0], get_true_format(test_row[10]))
        for i in range(1, 10):
            exec(f"row_201{i}_max[1] = max(row_201{i}_max[1], get_true_format(test_row[1{i}]))")
        row_2020_max[1] = max(row_2020_max[1], get_true_format(test_row[20]))
        for i in range(1, 10):
            exec(f"row_201{i}_max[2] = max(row_201{i}_max[2], get_true_format(test_row[2{i}]))")
        row_2020_max[2] = max(row_2020_max[2], get_true_format(test_row[30]))
        for i in range(1, 10):
            exec(f"row_201{i}_max[3] = max(row_201{i}_max[3], get_true_format(test_row[5{i}]))")
        row_2020_max[3] = max(row_2020_max[3], get_true_format(test_row[60]))
        for i in range(1, 10):
            exec(f"row_201{i}_max[5] = max(row_201{i}_max[5], get_true_format(test_row[8{i}]))")
        row_2020_max[5] = max(row_2020_max[5], get_true_format(test_row[90]))
        for i in range(1, 10):
            exec(f"row_201{i}_main[0] += get_true_format(test_row[{i}])")
        row_2020_main[0] += get_true_format(test_row[10])
        for i in range(1, 10):
            exec(f"row_201{i}_main[1] += get_true_format(test_row[1{i}])")
        row_2020_main[1] += get_true_format(test_row[20])
        for i in range(1, 10):
            exec(f"row_201{i}_main[2] += get_true_format(test_row[2{i}])")
        row_2020_main[2] += get_true_format(test_row[30])
        for i in range(1, 10):
            exec(f"row_201{i}_main[3] += get_true_format(test_row[5{i}])")
        row_2020_main[3] += get_true_format(test_row[60])
        for i in range(1, 10):
            exec(f"row_201{i}_main[5] += get_true_format(test_row[8{i}])")
        row_2020_main[5] += get_true_format(test_row[90])
    for i in range(1, 10):
        exec(f"row_201{i}_main[0] /= 3")
    row_2020_main[0] /= 3
    for i in range(1, 10):
        exec(f"row_201{i}_main[1] /= 3")
    row_2020_main[1] /= 3
    for i in range(1, 10):
        exec(f"row_201{i}_main[2] /= 3")
    row_2020_main[2] /= 3
    for i in range(1, 10):
        exec(f"row_201{i}_main[3] /= 3")
    row_2020_main[3] /= 3
    for i in range(1, 10):
        exec(f"row_201{i}_main[5] /= 3")
    row_2020_main[5] /= 3

    max_fire_data[2020].append(Fire(data_list[row_id][0].split()[0],
                               2011, row_2011_max[0], row_2011_max[1],
                               row_2011_max[2], row_2011_max[3], row_2011_max[4], row_2011_max[5]))
    max_fire_data[2019].append(Fire(data_list[row_id][0].split()[0],
                               2012, row_2012_max[0], row_2012_max[1],
                               row_2012_max[2], row_2012_max[3], row_2012_max[4], row_2012_max[5]))
    max_fire_data[2018].append(Fire(data_list[row_id][0].split()[0],
                               2013, row_2013_max[0], row_2013_max[1],
                               row_2013_max[2], row_2013_max[3], row_2013_max[4], row_2013_max[5]))
    max_fire_data[2017].append(Fire(data_list[row_id][0].split()[0],
                               2014, row_2014_max[0], row_2014_max[1],
                               row_2014_max[2], row_2014_max[3], row_2014_max[4], row_2014_max[5]))
    max_fire_data[2016].append(Fire(data_list[row_id][0].split()[0],
                               2015, row_2015_max[0], row_2015_max[1],
                               row_2015_max[2], row_2015_max[3], row_2015_max[4], row_2015_max[5]))
    max_fire_data[2015].append(Fire(data_list[row_id][0].split()[0],
                               2016, row_2015_max[0], row_2015_max[1],
                               row_2015_max[2], row_2015_max[3], row_2015_max[4], row_2015_max[5]))
    max_fire_data[2014].append(Fire(data_list[row_id][0].split()[0],
                               2017, row_2016_max[0], row_2016_max[1],
                               row_2016_max[2], row_2016_max[3], row_2016_max[4], row_2016_max[5]))
    max_fire_data[2013].append(Fire(data_list[row_id][0].split()[0],
                               2018, row_2018_max[0], row_2018_max[1],
                               row_2018_max[2], row_2018_max[3], row_2018_max[4], row_2018_max[5]))
    max_fire_data[2012].append(Fire(data_list[row_id][0].split()[0],
                               2019, row_2019_max[0], row_2019_max[1],
                               row_2019_max[2], row_2019_max[3], row_2019_max[4], row_2019_max[5]))
    max_fire_data[2011].append(Fire(data_list[row_id][0].split()[0],
                               2020, row_2020_max[0], row_2020_max[1],
                               row_2020_max[2], row_2020_max[3], row_2020_max[4], row_2020_max[5]))

    main_fire_data[2020].append(Fire(data_list[row_id][0].split()[0],
                                2011, row_2011_main[0], row_2011_main[1],
                                row_2011_main[2], row_2011_main[3], row_2011_main[4], row_2011_main[5]))
    main_fire_data[2019].append(Fire(data_list[row_id][0].split()[0],
                                2012, row_2012_main[0], row_2012_main[1],
                                row_2012_main[2], row_2012_main[3], row_2012_main[4], row_2012_main[5]))
    main_fire_data[2018].append(Fire(data_list[row_id][0].split()[0],
                                2013, row_2013_main[0], row_2013_main[1],
                                row_2013_main[2], row_2013_main[3], row_2013_main[4], row_2013_main[5]))
    main_fire_data[2017].append(Fire(data_list[row_id][0].split()[0],
                                2014, row_2014_main[0], row_2014_main[1],
                                row_2014_main[2], row_2014_main[3], row_2014_main[4], row_2014_main[5]))
    main_fire_data[2016].append(Fire(data_list[row_id][0].split()[0],
                                2015, row_2015_main[0], row_2015_main[1],
                                row_2015_main[2], row_2015_main[3], row_2015_main[4], row_2015_main[5]))
    main_fire_data[2015].append(Fire(data_list[row_id][0].split()[0],
                                2016, row_2015_main[0], row_2015_main[1],
                                row_2015_main[2], row_2015_main[3], row_2015_main[4], row_2015_main[5]))
    main_fire_data[2014].append(Fire(data_list[row_id][0].split()[0],
                                2017, row_2016_main[0], row_2016_main[1],
                                row_2016_main[2], row_2016_main[3], row_2016_main[4], row_2016_main[5]))
    main_fire_data[2013].append(Fire(data_list[row_id][0].split()[0],
                                2018, row_2018_main[0], row_2018_main[1],
                                row_2018_main[2], row_2018_main[3], row_2018_main[4], row_2018_main[5]))
    main_fire_data[2012].append(Fire(data_list[row_id][0].split()[0],
                                2019, row_2019_main[0], row_2019_main[1],
                                row_2019_main[2], row_2019_main[3], row_2019_main[4], row_2019_main[5]))
    main_fire_data[2011].append(Fire(data_list[row_id][0].split()[0],
                                2020, row_2020_main[0], row_2020_main[1],
                                row_2020_main[2], row_2020_main[3], row_2020_main[4], row_2020_main[5]))
work_book = openpyxl.Workbook()
for year in YEAR_LIST:
    work_sheet = work_book.create_sheet(title=str(year))
    work_sheet.append(["номер дня", 'atmospheric_pressure (средн)', 'precipitation_amount (средн)',
                       "relative_humidity (средн)", "soil_moisture (средн)", "temperature (средн)",
                       "soil_temperature (средн)", "dew_point (средн)", "komplex_kof (средн)", "mod_kof (средн)",
                       "mod_pok (средн)", 'atmospheric_pressure (макс)', 'precipitation_amount (макс)',
                       "relative_humidity (макс)", "soil_moisture (макс)", "temperature (макс)",
                       "soil_temperature (макс)", "dew_point (макс)", "komplex_kof (макс)",
                       "mod_kof (макс)", "mod_pok (макс)"])
    for fire_id in range(len(main_fire_data[year])):
        full_data_list = [main_fire_data[year][fire_id].get_full_fire_info()[0]]
        for fire_param in main_fire_data[year][fire_id].get_full_fire_info()[1:]:
            full_data_list.append(fire_param)
        for fire_param in max_fire_data[year][fire_id].get_full_fire_info()[1:]:
            full_data_list.append(fire_param)
        work_sheet.append(full_data_list)
work_book.save(f'{org_name}.xlsx')
work_book.close()

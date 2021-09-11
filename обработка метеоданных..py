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
YEAR_LIST = [2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020]


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
    def __init__(self, fire_day_number, fire_year, a_p_max, a_p_main, p_a_max, p_a_main,
                 r_h_max, r_h_main, s_m_max, s_m_main, t_max, t_main, s_t_max, s_t_main):
        self.fire_day_number = fire_day_number
        self.fire_year = fire_year
        self.atmospheric_pressure_max = get_true_format(a_p_max)
        self.atmospheric_pressure_main = get_true_format(a_p_main)
        self.precipitation_amount_max = get_true_format(p_a_max)
        self.precipitation_amount_main = get_true_format(p_a_main)
        self.relative_humidity_max = get_true_format(r_h_max)
        self.relative_humidity_main = get_true_format(r_h_main)
        self.soil_moisture_max = get_true_format(s_m_max)
        self.soil_moisture_main = get_true_format(s_m_main)
        self.temperature_max = get_true_format(t_max)
        self.temperature_main = get_true_format(t_main)
        self.soil_temperature_max = get_true_format(s_t_max)
        self.soil_temperature_main = get_true_format(s_t_main)
        self.dew_point = self.get_dew_point()
        self.komplex_kof = self.get_komplex_kof()
        self.komplex_pok = self.get_komplex_pok(main_fire_data)
        self.mod_kof = self.get_mod_kof()
        self.mod_pok = self.get_mod_pok(main_fire_data)

    def get_mod_kof(self):
        if self.precipitation_amount_main <= 0.5:
            return 1
        elif self.precipitation_amount_main <= 2:
            return 0.8
        elif self.precipitation_amount_main <= 5:
            return 0.4
        elif self.precipitation_amount_main <= 12:
            return 0.2
        elif self.precipitation_amount_main <= 19:
            return 0.1
        else:
            return 0

    def get_mod_pok(self, fire_data):
        if int(self.fire_day_number) == 2:
            return (self.temperature_main - self.dew_point) * self.temperature_main
        else:
            last_day = fire_data[self.fire_year][-1]
            return (last_day.mod_pok + (last_day.temperature_main - last_day.dew_point)
                    * last_day.temperature_main) * self.mod_kof

    def get_komplex_kof(self):
        if self.precipitation_amount_main < 3:
            return 0
        else:
            return 1

    def get_komplex_pok(self, fire_data):
        if int(self.fire_day_number) == 2:
            return (self.temperature_main - self.dew_point) * self.temperature_main
        else:
            last_day = fire_data[self.fire_year][-1]
            return (last_day.komplex_pok * self.komplex_kof + (self.temperature_main - self.dew_point) *
                    self.temperature_main)

    def get_dew_point(self):
        b = math.log10(self.relative_humidity_main / 4) + ((self.temperature_max / 4) * 17.27) /\
            (237.7 + (self.temperature_max / 4))
        return (237.7 * b) / (17.27 - b)

    def get_full_fire_info(self):
        return self.fire_day_number, self.atmospheric_pressure_max, self.atmospheric_pressure_main,\
               self.precipitation_amount_max, self.precipitation_amount_main, self.relative_humidity_max,\
               self.relative_humidity_main, self.soil_moisture_max, self.soil_moisture_main, self.temperature_max,\
               self.temperature_main, self.soil_temperature_max, self.soil_temperature_main, self.dew_point,\
               self.komplex_kof, self.komplex_pok, self.mod_kof, self.mod_pok


flag1 = True
flag2 = True
flag3 = True

test = ['DATA_20210820_125109.csv', 'DATA_20210820_125111.csv', 'DATA_20210820_125114.csv', 'DATA_20210820_125117.csv']

list_files = ['DATA_20210831_193317.csv']
for fli in list_files:
    with open(fli, encoding="utf8") as f:
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
    for row_id in range(8, len(data_list) - 168, 8):
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
        for test_row in data_list[row_id-1:row_id+7]:
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
                if len(test_row) != 73:
                    exec(f"row_201{i}_max[5] = max(row_201{i}_max[5], get_true_format(test_row[8{i}]))")
            if len(test_row) != 73:
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
            if len(test_row) != 73:
                for i in range(1, 10):
                    exec(f"row_201{i}_main[5] += get_true_format(test_row[8{i}])")
                row_2020_main[5] += get_true_format(test_row[90])
        for i in range(1, 10):
            exec(f"row_201{i}_main[0] /= 4")
        row_2020_main[0] /= 4
        for i in range(1, 10):
            exec(f"row_201{i}_main[1] /= 4")
        row_2020_main[1] /= 4
        for i in range(1, 10):
            exec(f"row_201{i}_main[2] /= 4")
        row_2020_main[2] /= 4
        for i in range(1, 10):
            exec(f"row_201{i}_main[3] /= 4")
        row_2020_main[3] /= 4
        for i in range(1, 10):
            exec(f"row_201{i}_main[5] /= 4")
        row_2020_main[5] /= 4
        main_fire_data[2020].append(Fire(data_list[row_id][0].split()[0], 2011, row_2011_max[0],
                                         row_2011_main[0], row_2011_max[1], row_2011_main[1],
                                         row_2011_max[2], row_2011_main[2],
                                         row_2011_max[3], row_2011_main[3],
                                         row_2011_max[4], row_2011_main[4],
                                         row_2011_max[5], row_2011_main[5]))
        main_fire_data[2019].append(Fire(data_list[row_id][0].split()[0], 2012, row_2012_max[0],
                                         row_2012_main[0], row_2012_max[1],
                                         row_2012_main[1],
                                         row_2012_max[2], row_2012_main[2],
                                         row_2012_max[3], row_2012_main[3],
                                         row_2012_max[4], row_2012_main[4],
                                         row_2012_max[5], row_2012_main[5]))
        main_fire_data[2018].append(Fire(data_list[row_id][0].split()[0], 2013, row_2013_max[0],
                                         row_2013_main[0], row_2013_max[1], row_2013_main[1],
                                         row_2013_max[2], row_2013_main[2],
                                         row_2013_max[3], row_2013_main[3],
                                         row_2013_max[4], row_2013_main[4],
                                         row_2013_max[5], row_2013_main[5]))
        main_fire_data[2017].append(Fire(data_list[row_id][0].split()[0], 2014, row_2014_max[0],
                                         row_2014_main[0], row_2014_max[1], row_2014_main[1],
                                         row_2014_max[2], row_2014_main[2],
                                         row_2014_max[3], row_2014_main[3],
                                         row_2014_max[4], row_2014_main[4],
                                         row_2014_max[5], row_2014_main[5]))
        main_fire_data[2016].append(Fire(data_list[row_id][0].split()[0], 2015, row_2015_max[0],
                                         row_2015_main[0], row_2015_max[1], row_2015_main[1],
                                         row_2015_max[2], row_2015_main[2],
                                         row_2015_max[3], row_2015_main[3],
                                         row_2015_max[4], row_2015_main[4],
                                         row_2015_max[5], row_2015_main[5]))
        main_fire_data[2015].append(Fire(data_list[row_id][0].split()[0], 2016, row_2016_max[0],
                                         row_2016_main[0], row_2016_max[1], row_2016_main[1],
                                         row_2016_max[2], row_2016_main[2],
                                         row_2016_max[3], row_2016_main[3],
                                         row_2016_max[4], row_2016_main[4],
                                         row_2016_max[5], row_2016_main[5]))
        main_fire_data[2014].append(Fire(data_list[row_id][0].split()[0], 2017, row_2017_max[0],
                                         row_2017_main[0], row_2017_max[1], row_2017_main[1],
                                         row_2017_max[2], row_2017_main[2],
                                         row_2017_max[3], row_2017_main[3],
                                         row_2017_max[4], row_2017_main[4],
                                         row_2017_max[5], row_2017_main[5]))
        main_fire_data[2013].append(Fire(data_list[row_id][0].split()[0], 2018, row_2018_max[0],
                                         row_2018_main[0], row_2018_max[1], row_2018_main[1],
                                         row_2018_max[2], row_2018_main[2],
                                         row_2018_max[3], row_2018_main[3],
                                         row_2018_max[4], row_2018_main[4],
                                         row_2018_max[5], row_2018_main[5]))
        main_fire_data[2012].append(Fire(data_list[row_id][0].split()[0], 2019, row_2019_max[0],
                                         row_2019_main[0], row_2019_max[1], row_2019_main[1],
                                         row_2019_max[2], row_2019_main[2],
                                         row_2019_max[3], row_2019_main[3],
                                         row_2019_max[4], row_2019_main[4],
                                         row_2019_max[5], row_2019_main[5]))
        main_fire_data[2011].append(Fire(data_list[row_id][0].split()[0], 2020, row_2020_max[0],
                                         row_2020_main[0], row_2020_max[1], row_2020_main[1],
                                         row_2020_max[2], row_2020_main[2],
                                         row_2020_max[3], row_2020_main[3],
                                         row_2020_max[4], row_2020_main[4],
                                         row_2020_max[5], row_2020_main[5]))
    work_book = openpyxl.Workbook()
    for year in YEAR_LIST:
        work_sheet = work_book.create_sheet(title=str(YEAR_LIST[len(YEAR_LIST) - 1 - YEAR_LIST.index(year)]))
        work_sheet.append(["номер дня", "atmospheric_pressure_max", "atmospheric_pressure_main",
                           "precipitation_amount_max", 'precipitation_amount_main', 'relative_humidity_max',
                           'relative_humidity_main', 'soil_moisture_max', 'soil_moisture_main', 'temperature_max',
                           'temperature_main', 'soil_temperature_max', 'soil_temperature_main', 'dew_point',
                           'komplex_kof', 'komplex_pok', 'mod_kof', 'mod_pok'])
        for fire_id in range(len(main_fire_data[year])):
            full_data_list = [main_fire_data[year][fire_id].get_full_fire_info()[0]]
            for fire_param in main_fire_data[year][fire_id].get_full_fire_info()[1:]:
                full_data_list.append(fire_param)
            work_sheet.append(full_data_list)
    work_book.save(f'{org_name}.xlsx')
    work_book.close()

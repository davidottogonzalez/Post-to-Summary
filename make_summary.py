from openpyxl import load_workbook
import shutil, os, glob, collections

source_wb = None
summary_wb = None


def setup(source_filename, equiv):
    if equiv:
        new_template = source_filename.replace("atp_Q2_new", "summary_equiv")
    else:
        new_template = source_filename.replace("atp_Q2_new", "summary_unequiv")
    shutil.copy('template.xlsx', new_template)

    global source_wb
    global summary_wb

    source_wb = load_workbook('./preprocessed/' + source_filename)
    summary_wb = load_workbook(new_template)

    return new_template


def process_summary_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_sheet = source_wb.get_sheet_by_name('Summary')
    source_spot_sheet = source_wb.get_sheet_by_name('Spot Detail')
    summary_sheet = summary_wb.get_sheet_by_name('Summary Metrics')
    program_metrics_sheet = summary_wb.get_sheet_by_name('Program Metrics')
    start_row = 0

    summary_sheet.cell(row=summary_sheet.max_row + 2, column=1).value = 'Network Summary'

    for row_num in range(1, source_sheet.max_row + 1):
        if start_row == 0:
            start_row = summary_sheet.max_row + 2
        else:
            start_row = summary_sheet.max_row + 1

        for col_num in range(1, source_sheet.max_column + 1):
            summary_sheet.cell(row=start_row, column=col_num).value = source_sheet.cell(row=row_num,
                                                                                        column=col_num).value

    start_row = 0
    p_start_row = 1

    summary_sheet.cell(row=summary_sheet.max_row + 2, column=1).value = 'Spot Detail'

    for row_num in range(1, source_spot_sheet.max_row + 1):
        if start_row == 0:
            start_row = summary_sheet.max_row + 2
        else:
            start_row = summary_sheet.max_row + 1

        for col_num in range(1, source_spot_sheet.max_column + 1):
            summary_sheet.cell(row=start_row, column=col_num).value = source_spot_sheet.cell(row=row_num,
                                                                                             column=col_num).value
            program_metrics_sheet.cell(row=p_start_row, column=col_num).value = source_spot_sheet.cell(row=row_num,
                                                                                                       column=col_num).value

        p_start_row += 1

    # start totals
    for row_num in range(1, source_sheet.max_row + 1):
        if source_sheet.cell(row=row_num, column=1).value == 'Total':
            for col_num in range(1, source_sheet.max_column + 1):
                if source_sheet.cell(row=1, column=col_num).value == 'total_impressions' and not equiv:
                    summary_sheet.cell(row=6, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_unequiv_impressions' and equiv:
                    summary_sheet.cell(row=6, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_reach':
                    summary_sheet.cell(row=8, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_reach_pct':
                    summary_sheet.cell(row=10, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_effective_reach':
                    summary_sheet.cell(row=12, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_effective_reach_pct':
                    summary_sheet.cell(row=13, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'total_reach_raw_count':
                    summary_sheet.cell(row=9, column=3).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_impressions' and not equiv:
                    summary_sheet.cell(row=6, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_unequiv_impressions' and equiv:
                    summary_sheet.cell(row=6, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_reach':
                    summary_sheet.cell(row=8, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_reach_pct':
                    summary_sheet.cell(row=10, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_effective_reach':
                    summary_sheet.cell(row=12, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_effective_reach_pct':
                    summary_sheet.cell(row=13, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_reach_raw_count':
                    summary_sheet.cell(row=9, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_index_impressions' and not equiv:
                    summary_sheet.cell(row=7, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_index_unequiv_impressions' and equiv:
                    summary_sheet.cell(row=7, column=2).value = source_sheet.cell(row=row_num, column=col_num).value
                if source_sheet.cell(row=1, column=col_num).value == 'target_index_reach':
                    summary_sheet.cell(row=11, column=2).value = source_sheet.cell(row=row_num, column=col_num).value

            # insert 1 for indexes of total
            summary_sheet.cell(row=7, column=3).value = 100.0
            summary_sheet.cell(row=11, column=3).value = 100.0
            summary_sheet.cell(row=14, column=3).value = 100.0

            # custom calculations
            # Index Effective Reach
            summary_sheet.cell(row=14, column=2).value = summary_sheet.cell(row=13, column=2).value / summary_sheet.cell(row=13, column=3).value if summary_sheet.cell(row=13, column=2).value else 0
            # Average Frequency
            summary_sheet.cell(row=15, column=2).value = summary_sheet.cell(row=6, column=2).value / summary_sheet.cell(row=8, column=2).value
            summary_sheet.cell(row=15, column=3).value = summary_sheet.cell(row=6, column=3).value / summary_sheet.cell(row=8, column=3).value
            # Avg Freq % Diff v. Total
            summary_sheet.cell(row=16, column=2).value = (summary_sheet.cell(row=15, column=2).value - summary_sheet.cell(row=15, column=3).value) / summary_sheet.cell(row=15, column=3).value
            summary_sheet.cell(row=16, column=3).value = (summary_sheet.cell(row=15, column=3).value - summary_sheet.cell(row=15, column=3).value) / summary_sheet.cell(row=15, column=3).value

            # Projected Calculations
            summary_sheet.cell(row=4, column=2).value = summary_sheet.cell(row=8, column=2).value / summary_sheet.cell(row=10, column=2).value
            summary_sheet.cell(row=4, column=3).value = summary_sheet.cell(row=8, column=3).value / summary_sheet.cell(row=10, column=3).value
            summary_sheet.cell(row=5, column=2).value = summary_sheet.cell(row=4, column=2).value / summary_sheet.cell(row=4, column=3).value
            summary_sheet.cell(row=5, column=3).value = summary_sheet.cell(row=4, column=3).value / summary_sheet.cell(row=4, column=3).value

            break

    summary_wb.save(filename)

    return True


def process_Network_Daypart_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_network_day_sheet = source_wb.get_sheet_by_name('Network Daypart')
    dest_network_day_sheet = summary_wb.create_sheet("Network Daypart", 2)

    for row_num in range(1, source_network_day_sheet.max_row + 1):
        for col_num in range(1, source_network_day_sheet.max_column + 1):
            dest_network_day_sheet.cell(row=row_num, column=col_num).value = source_network_day_sheet.cell(row=row_num, column=col_num).value

    summary_wb.save(filename)

    return True


def process_frequency_distribution_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_freq_sheet = source_wb.get_sheet_by_name('Frequency Distribution')
    dest_freq_sheet = summary_wb.get_sheet_by_name('Frequency Distribution')
    dest_row = 2
    source_rows = source_freq_sheet.max_row

    for row_num in range(1, source_rows + 1):
        if source_freq_sheet.cell(row=row_num, column=1).value == 'Spot' and source_freq_sheet.cell(row=row_num, column=2).value == 'Total':
            for col_num in range(1, source_freq_sheet.max_column + 1):
                if source_freq_sheet.cell(row=1, column=col_num).value == 'frequency':
                    dest_freq_sheet.cell(row=dest_row, column=1).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'target':
                    dest_freq_sheet.cell(row=dest_row, column=3).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'total':
                    dest_freq_sheet.cell(row=dest_row, column=2).value = source_freq_sheet.cell(row=row_num, column=col_num).value
            dest_row += 1

    dest_row += 1
    dest_freq_sheet.cell(row=dest_row, column=1).value = '# Networks'
    dest_freq_sheet.cell(row=dest_row, column=2).value = 'Total (Campaign)'
    dest_freq_sheet.cell(row=dest_row, column=3).value = 'Target (Campaign)'
    dest_freq_sheet.cell(row=dest_row, column=4).value = 'Target Composition'
    dest_row += 1

    for row_num in range(1, source_rows + 1):
        if source_freq_sheet.cell(row=row_num, column=1).value == 'Network' and source_freq_sheet.cell(row=row_num, column=2).value == 'Total':
            for col_num in range(1, source_freq_sheet.max_column + 1):
                if source_freq_sheet.cell(row=1, column=col_num).value == 'frequency':
                    dest_freq_sheet.cell(row=dest_row, column=1).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'target':
                    dest_freq_sheet.cell(row=dest_row, column=3).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'total':
                    dest_freq_sheet.cell(row=dest_row, column=2).value = source_freq_sheet.cell(row=row_num, column=col_num).value
            dest_row += 1

    dest_freq_sheet.cell(row=dest_row, column=1).value = 'Total'
    dest_row += 2
    dest_freq_sheet.cell(row=dest_row, column=1).value = '# Programs'
    dest_freq_sheet.cell(row=dest_row, column=2).value = 'Total (Campaign)'
    dest_freq_sheet.cell(row=dest_row, column=3).value = 'Target (Campaign)'
    dest_row += 1

    for row_num in range(1, source_rows + 1):
        if source_freq_sheet.cell(row=row_num, column=1).value == 'Program' and source_freq_sheet.cell(row=row_num, column=2).value == 'Total':
            for col_num in range(1, source_freq_sheet.max_column + 1):
                if source_freq_sheet.cell(row=1, column=col_num).value == 'frequency':
                    dest_freq_sheet.cell(row=dest_row, column=1).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'target':
                    dest_freq_sheet.cell(row=dest_row, column=3).value = source_freq_sheet.cell(row=row_num, column=col_num).value
                if source_freq_sheet.cell(row=1, column=col_num).value == 'total':
                    dest_freq_sheet.cell(row=dest_row, column=2).value = source_freq_sheet.cell(row=row_num, column=col_num).value
            dest_row += 1

    # Calculations
    incrementor = 1
    sum_incrementor = 0
    network_start = 0
    for row_num in range(2, dest_freq_sheet.max_row + 1):
        if incrementor % 5 == 0 or not dest_freq_sheet.cell(row=row_num + 1, column=1).value:
            if dest_freq_sheet.cell(row=row_num,column=3).value:
                dest_freq_sheet.cell(row=row_num,column=4).value = sum_incrementor + dest_freq_sheet.cell(row=row_num,column=3).value
            else:
                dest_freq_sheet.cell(row=row_num,column=4).value = sum_incrementor
            incrementor = 1
            sum_incrementor = 0
        else:
            if dest_freq_sheet.cell(row=row_num,column=3).value:
                sum_incrementor += dest_freq_sheet.cell(row=row_num,column=3).value
            incrementor += 1
        if not dest_freq_sheet.cell(row=row_num + 1, column=1).value:
            network_start = row_num + 3
            break

    network_total = 0
    network_target = 0
    for row_num in range(network_start, dest_freq_sheet.max_row + 1):
        if dest_freq_sheet.cell(row=row_num, column=1).value != 'Total':
            network_total += dest_freq_sheet.cell(row=row_num, column=2).value
            network_target += dest_freq_sheet.cell(row=row_num, column=3).value if dest_freq_sheet.cell(row=row_num, column=3).value else 0
        else:
            dest_freq_sheet.cell(row=row_num, column=2).value = network_total
            dest_freq_sheet.cell(row=row_num, column=3).value = network_target
            break

    for row_num in range(network_start, dest_freq_sheet.max_row + 1):
        if dest_freq_sheet.cell(row=row_num, column=1).value != 'Total':
            dest_freq_sheet.cell(row=row_num, column=4).value = dest_freq_sheet.cell(row=row_num, column=3).value / network_target if dest_freq_sheet.cell(row=row_num, column=3).value else 0
        else:
            break

    summary_wb.save(filename)

    return True


def process_reach_by_week_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_reach_by_week_sheet = source_wb.get_sheet_by_name('Reach by Week')
    dest_reach_by_week_sheet = summary_wb.get_sheet_by_name('Reach by Week')
    dest_row = 2
    write_column = 1

    equiv_list = [6,11]
    unequiv_list = [7,12]

    for row_num in range(2, source_reach_by_week_sheet.max_row + 1):
        if source_reach_by_week_sheet.cell(row=row_num, column=1).value == 'Total':
            for col_num in range(2, source_reach_by_week_sheet.max_column + 1):
                if not equiv and col_num in equiv_list:
                    continue
                if equiv and col_num in unequiv_list:
                    continue

                dest_reach_by_week_sheet.cell(row=dest_row, column=write_column).value = source_reach_by_week_sheet.cell(row=row_num, column=col_num).value
                write_column += 1
            dest_row += 1
            write_column = 1

    summary_wb.save(filename)
    return True


def process_frequency_distribution_by_net_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_freq_sheet = source_wb.get_sheet_by_name('Frequency Distribution')
    dest_freq_sheet = summary_wb.get_sheet_by_name('Freq Distribution by Net')
    source_rows = source_freq_sheet.max_row
    freq_obj = {}

    for row_num in range(2, source_rows + 1):
        if freq_obj.has_key(int(source_freq_sheet.cell(row=row_num, column=3).value)):
            if freq_obj[int(source_freq_sheet.cell(row=row_num, column=3).value)].has_key(source_freq_sheet.cell(row=row_num, column=1).value):
                freq_obj[int(source_freq_sheet.cell(row=row_num, column=3).value)][source_freq_sheet.cell(row=row_num, column=1).value].append(
                    {source_freq_sheet.cell(row=row_num, column=2).value:
                    {'target':source_freq_sheet.cell(row=row_num, column=4).value,
                     'total':source_freq_sheet.cell(row=row_num, column=5).value}}
                )
            else:
                freq_obj[int(source_freq_sheet.cell(row=row_num, column=3).value)][source_freq_sheet.cell(row=row_num, column=1).value] = [
                    {source_freq_sheet.cell(row=row_num, column=2).value:
                    {'target':source_freq_sheet.cell(row=row_num, column=4).value,
                     'total':source_freq_sheet.cell(row=row_num, column=5).value}}
                ]
        else:
            freq_obj[int(source_freq_sheet.cell(row=row_num, column=3).value)] = {source_freq_sheet.cell(row=row_num, column=1).value:[{
                source_freq_sheet.cell(row=row_num, column=2).value:
                    {'target':source_freq_sheet.cell(row=row_num, column=4).value,
                     'total':source_freq_sheet.cell(row=row_num, column=5).value}}]}

    write_row = 2
    for freq, row_list in collections.OrderedDict(sorted(freq_obj.items())).items():
        dest_freq_sheet.cell(row=write_row,column=1).value = freq
        for row_obj in row_list['Spot']:
            if row_obj.has_key('Total'):
                dest_freq_sheet.cell(row=write_row, column=2).value = row_obj['Total']['total']
                dest_freq_sheet.cell(row=write_row, column=3).value = row_obj['Total']['target']
            if row_obj.has_key('Bravo'):
                dest_freq_sheet.cell(row=write_row, column=4).value = row_obj['Bravo']['total']
                dest_freq_sheet.cell(row=write_row, column=5).value = row_obj['Bravo']['target']
            if row_obj.has_key('CNBC'):
                dest_freq_sheet.cell(row=write_row, column=6).value = row_obj['CNBC']['total']
                dest_freq_sheet.cell(row=write_row, column=7).value = row_obj['CNBC']['target']
            if row_obj.has_key('E!'):
                dest_freq_sheet.cell(row=write_row, column=8).value = row_obj['E!']['total']
                dest_freq_sheet.cell(row=write_row, column=9).value = row_obj['E!']['target']
            if row_obj.has_key('Golf Channel'):
                dest_freq_sheet.cell(row=write_row, column=10).value = row_obj['Golf Channel']['total']
                dest_freq_sheet.cell(row=write_row, column=11).value = row_obj['Golf Channel']['target']
            if row_obj.has_key('MSNBC'):
                dest_freq_sheet.cell(row=write_row, column=12).value = row_obj['MSNBC']['total']
                dest_freq_sheet.cell(row=write_row, column=13).value = row_obj['MSNBC']['target']
            if row_obj.has_key('NBCSN'):
                dest_freq_sheet.cell(row=write_row, column=14).value = row_obj['NBCSN']['total']
                dest_freq_sheet.cell(row=write_row, column=15).value = row_obj['NBCSN']['target']
            if row_obj.has_key('Oxygen'):
                dest_freq_sheet.cell(row=write_row, column=16).value = row_obj['Oxygen']['total']
                dest_freq_sheet.cell(row=write_row, column=17).value = row_obj['Oxygen']['target']
            if row_obj.has_key('Syfy'):
                dest_freq_sheet.cell(row=write_row, column=18).value = row_obj['Syfy']['total']
                dest_freq_sheet.cell(row=write_row, column=19).value = row_obj['Syfy']['target']
            if row_obj.has_key('USA'):
                dest_freq_sheet.cell(row=write_row, column=20).value = row_obj['USA']['total']
                dest_freq_sheet.cell(row=write_row, column=21).value = row_obj['USA']['target']
            if row_obj.has_key('NBC'):
                dest_freq_sheet.cell(row=write_row, column=22).value = row_obj['NBC']['total']
                dest_freq_sheet.cell(row=write_row, column=23).value = row_obj['NBC']['target']
            if row_obj.has_key('Esquire'):
                dest_freq_sheet.cell(row=write_row, column=24).value = row_obj['Esquire']['total']
                dest_freq_sheet.cell(row=write_row, column=25).value = row_obj['Esquire']['target']
        write_row += 1

    write_row += 1
    dest_freq_sheet.cell(row=write_row, column=1).value = '# Programs'
    for col_num in range(2, dest_freq_sheet.max_column + 1):
        dest_freq_sheet.cell(row=write_row, column=col_num).value = dest_freq_sheet.cell(row=1, column=col_num).value
    write_row += 1

    for freq, row_list in collections.OrderedDict(freq_obj.items()).items():
        dest_freq_sheet.cell(row=write_row,column=1).value = freq
        if row_list.has_key('Program'):
            for row_obj in row_list['Program']:
                if row_obj.has_key('Total'):
                    dest_freq_sheet.cell(row=write_row, column=2).value = row_obj['Total']['total']
                    dest_freq_sheet.cell(row=write_row, column=3).value = row_obj['Total']['target']
                if row_obj.has_key('Bravo'):
                    dest_freq_sheet.cell(row=write_row, column=4).value = row_obj['Bravo']['total']
                    dest_freq_sheet.cell(row=write_row, column=5).value = row_obj['Bravo']['target']
                if row_obj.has_key('CNBC'):
                    dest_freq_sheet.cell(row=write_row, column=6).value = row_obj['CNBC']['total']
                    dest_freq_sheet.cell(row=write_row, column=7).value = row_obj['CNBC']['target']
                if row_obj.has_key('E!'):
                    dest_freq_sheet.cell(row=write_row, column=8).value = row_obj['E!']['total']
                    dest_freq_sheet.cell(row=write_row, column=9).value = row_obj['E!']['target']
                if row_obj.has_key('Golf Channel'):
                    dest_freq_sheet.cell(row=write_row, column=10).value = row_obj['Golf Channel']['total']
                    dest_freq_sheet.cell(row=write_row, column=11).value = row_obj['Golf Channel']['target']
                if row_obj.has_key('MSNBC'):
                    dest_freq_sheet.cell(row=write_row, column=12).value = row_obj['MSNBC']['total']
                    dest_freq_sheet.cell(row=write_row, column=13).value = row_obj['MSNBC']['target']
                if row_obj.has_key('NBCSN'):
                    dest_freq_sheet.cell(row=write_row, column=13).value = row_obj['NBCSN']['total']
                    dest_freq_sheet.cell(row=write_row, column=14).value = row_obj['NBCSN']['target']
                if row_obj.has_key('Oxygen'):
                    dest_freq_sheet.cell(row=write_row, column=15).value = row_obj['Oxygen']['total']
                    dest_freq_sheet.cell(row=write_row, column=16).value = row_obj['Oxygen']['target']
                if row_obj.has_key('Syfy'):
                    dest_freq_sheet.cell(row=write_row, column=17).value = row_obj['Syfy']['total']
                    dest_freq_sheet.cell(row=write_row, column=18).value = row_obj['Syfy']['target']
                if row_obj.has_key('USA'):
                    dest_freq_sheet.cell(row=write_row, column=19).value = row_obj['USA']['total']
                    dest_freq_sheet.cell(row=write_row, column=20).value = row_obj['USA']['target']
                if row_obj.has_key('NBC'):
                    dest_freq_sheet.cell(row=write_row, column=21).value = row_obj['NBC']['total']
                    dest_freq_sheet.cell(row=write_row, column=22).value = row_obj['NBC']['target']
                if row_obj.has_key('Esquire'):
                    dest_freq_sheet.cell(row=write_row, column=23).value = row_obj['Esquire']['total']
                    dest_freq_sheet.cell(row=write_row, column=24).value = row_obj['Esquire']['target']
            write_row += 1

    summary_wb.save(filename)

    return True


def process_network_reach_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_reach_net_sheet = source_wb.get_sheet_by_name('Reach by Week')
    dest_reach_net_sheet = summary_wb.get_sheet_by_name('Network Reach by Week')
    source_rows = source_reach_net_sheet.max_row
    reach_net_obj = {}

    for row_num in range(2, source_rows + 1):
        if reach_net_obj.has_key(int(source_reach_net_sheet.cell(row=row_num, column=2).value)):
            reach_net_obj[int(source_reach_net_sheet.cell(row=row_num, column=2).value)][source_reach_net_sheet.cell(row=row_num, column=1).value] =\
                {'weekof':source_reach_net_sheet.cell(row=row_num, column=3).value,
                 'total':source_reach_net_sheet.cell(row=row_num, column=4).value,
                 'total_pct':source_reach_net_sheet.cell(row=row_num, column=5).value,
                 'total_impressions':source_reach_net_sheet.cell(row=row_num, column=6).value,
                 'total_impressions_unequiv':source_reach_net_sheet.cell(row=row_num, column=7).value,
                 'total_frequency_unequiv':source_reach_net_sheet.cell(row=row_num, column=8).value,
                 'target':source_reach_net_sheet.cell(row=row_num, column=9).value,
                 'target_pct':source_reach_net_sheet.cell(row=row_num, column=10).value,
                 'target_impressions':source_reach_net_sheet.cell(row=row_num, column=11).value,
                 'target_impressions_unequiv':source_reach_net_sheet.cell(row=row_num, column=12).value,
                 'target_frequency_unequiv':source_reach_net_sheet.cell(row=row_num, column=13).value}
        else:
            reach_net_obj[int(source_reach_net_sheet.cell(row=row_num, column=2).value)] = {source_reach_net_sheet.cell(row=row_num, column=1).value:
                                                                                       {'weekof':source_reach_net_sheet.cell(row=row_num, column=3).value,
                                                                                        'total':source_reach_net_sheet.cell(row=row_num, column=4).value,
                                                                                        'total_pct':source_reach_net_sheet.cell(row=row_num, column=5).value,
                                                                                        'total_impressions':source_reach_net_sheet.cell(row=row_num, column=6).value,
                                                                                        'total_impressions_unequiv':source_reach_net_sheet.cell(row=row_num, column=7).value,
                                                                                        'total_frequency_unequiv':source_reach_net_sheet.cell(row=row_num, column=8).value,
                                                                                        'target':source_reach_net_sheet.cell(row=row_num, column=9).value,
                                                                                        'target_pct':source_reach_net_sheet.cell(row=row_num, column=10).value,
                                                                                        'target_impressions':source_reach_net_sheet.cell(row=row_num, column=11).value,
                                                                                        'target_impressions_unequiv':source_reach_net_sheet.cell(row=row_num, column=12).value,
                                                                                        'target_frequency_unequiv':source_reach_net_sheet.cell(row=row_num, column=13).value}}
    write_row = 2
    for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
        dest_reach_net_sheet.cell(row=write_row,column=1).value = row_list['Total']['weekof']
        dest_reach_net_sheet.cell(row=write_row,column=2).value = row_list['Total']['target_pct']
        dest_reach_net_sheet.cell(row=write_row,column=3).value = row_list['NBC']['target_pct'] if row_list.has_key('NBC') else ''
        dest_reach_net_sheet.cell(row=write_row,column=5).value = row_list['Total']['weekof']
        dest_reach_net_sheet.cell(row=write_row,column=6).value = row_list['Bravo']['target_pct'] if row_list.has_key('Bravo') else ''
        dest_reach_net_sheet.cell(row=write_row,column=7).value = row_list['CNBC']['target_pct'] if row_list.has_key('CNBC') else ''
        dest_reach_net_sheet.cell(row=write_row,column=8).value = row_list['Chiller']['target_pct'] if row_list.has_key('Chiller') else ''
        dest_reach_net_sheet.cell(row=write_row,column=9).value = row_list['E!']['target_pct'] if row_list.has_key('E!') else ''
        dest_reach_net_sheet.cell(row=write_row,column=10).value = row_list['Esquire']['target_pct'] if row_list.has_key('Esquire') else ''
        dest_reach_net_sheet.cell(row=write_row,column=11).value = row_list['Golf Channel']['target_pct'] if row_list.has_key('Golf Channel') else ''
        dest_reach_net_sheet.cell(row=write_row,column=12).value = row_list['MSNBC']['target_pct'] if row_list.has_key('MSNBC') else ''
        dest_reach_net_sheet.cell(row=write_row,column=13).value = row_list['NBCSN']['target_pct'] if row_list.has_key('NBCSN') else ''
        dest_reach_net_sheet.cell(row=write_row,column=14).value = row_list['Oxygen']['target_pct'] if row_list.has_key('Oxygen') else ''
        dest_reach_net_sheet.cell(row=write_row,column=15).value = row_list['Syfy']['target_pct'] if row_list.has_key('Syfy') else ''
        dest_reach_net_sheet.cell(row=write_row,column=16).value = row_list['USA']['target_pct'] if row_list.has_key('USA') else ''
        write_row += 1

    # Bravo
    if row_list.has_key('Bravo'):
        write_row += 3
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Bravo'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Bravo']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Bravo']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Bravo']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Bravo']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Bravo']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Bravo']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Bravo']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Bravo']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Bravo']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Bravo']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Bravo']['target_frequency_unequiv']
            write_row += 1

    # CNBC
    if row_list.has_key('CNBC'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'CNBC'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['CNBC']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['CNBC']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['CNBC']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['CNBC']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['CNBC']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['CNBC']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['CNBC']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['CNBC']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['CNBC']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['CNBC']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['CNBC']['target_frequency_unequiv']
            write_row += 1

    # Chiller
    if row_list.has_key('Chiller'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Chiller'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Chiller']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Chiller']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Chiller']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Chiller']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Chiller']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Chiller']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Chiller']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Chiller']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Chiller']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Chiller']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Chiller']['target_frequency_unequiv']
            write_row += 1

    # E!
    if row_list.has_key('E!'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'E!'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['E!']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['E!']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['E!']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['E!']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['E!']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['E!']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['E!']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['E!']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['E!']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['E!']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['E!']['target_frequency_unequiv']
            write_row += 1

    # Esquire
    if row_list.has_key('Esquire'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Esquire'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Esquire']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Esquire']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Esquire']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Esquire']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Esquire']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Esquire']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Esquire']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Esquire']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Esquire']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Esquire']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Esquire']['target_frequency_unequiv']
            write_row += 1

    # Golf Channel
    if row_list.has_key('Golf Channel'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Golf Channel'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Golf Channel']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Golf Channel']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Golf Channel']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Golf Channel']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Golf Channel']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Golf Channel']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Golf Channel']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Golf Channel']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Golf Channel']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Golf Channel']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Golf Channel']['target_frequency_unequiv']
            write_row += 1

    # MSNBC
    if row_list.has_key('MSNBC'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'MSNBC'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['MSNBC']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['MSNBC']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['MSNBC']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['MSNBC']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['MSNBC']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['MSNBC']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['MSNBC']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['MSNBC']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['MSNBC']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['MSNBC']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['MSNBC']['target_frequency_unequiv']
            write_row += 1

    # NBC
    if row_list.has_key('NBC'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'NBC'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['NBC']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['NBC']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['NBC']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['NBC']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['NBC']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['NBC']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['NBC']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['NBC']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['NBC']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['NBC']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['NBC']['target_frequency_unequiv']
            write_row += 1

    # NBCSN
    if row_list.has_key('NBCSN'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'NBCSN'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['NBCSN']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['NBCSN']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['NBCSN']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['NBCSN']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['NBCSN']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['NBCSN']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['NBCSN']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['NBCSN']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['NBCSN']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['NBCSN']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['NBCSN']['target_frequency_unequiv']
            write_row += 1

    # Oxygen
    if row_list.has_key('Oxygen'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Oxygen'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Oxygen']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Oxygen']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Oxygen']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Oxygen']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Oxygen']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Oxygen']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Oxygen']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Oxygen']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Oxygen']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Oxygen']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Oxygen']['target_frequency_unequiv']
            write_row += 1

    # Syfy
    if row_list.has_key('Syfy'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'Syfy'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['Syfy']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['Syfy']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['Syfy']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Syfy']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['Syfy']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['Syfy']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['Syfy']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['Syfy']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Syfy']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['Syfy']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['Syfy']['target_frequency_unequiv']
            write_row += 1

    # USA
    if row_list.has_key('USA'):
        write_row += 2
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'USA'
        write_row += 1
        dest_reach_net_sheet.cell(row=write_row, column=1).value = 'week'
        dest_reach_net_sheet.cell(row=write_row, column=2).value = 'week_of'
        dest_reach_net_sheet.cell(row=write_row, column=3).value = 'reach_total'
        dest_reach_net_sheet.cell(row=write_row, column=4).value = 'reach_pct_total'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_total'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=5).value = 'impressions_unequiv_total'
        dest_reach_net_sheet.cell(row=write_row, column=6).value = 'avg_freq_total'
        dest_reach_net_sheet.cell(row=write_row, column=7).value = 'reach_target'
        dest_reach_net_sheet.cell(row=write_row, column=8).value = 'reach_pct_target'
        if equiv:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_target'
        else:
            dest_reach_net_sheet.cell(row=write_row, column=9).value = 'impressions_unequiv_target'
        dest_reach_net_sheet.cell(row=write_row, column=10).value = 'avg_freq_target'
        write_row += 1

        for week, row_list in collections.OrderedDict(sorted(reach_net_obj.items())).items():
            dest_reach_net_sheet.cell(row=write_row, column=1).value = week
            dest_reach_net_sheet.cell(row=write_row, column=2).value = row_list['USA']['weekof']
            dest_reach_net_sheet.cell(row=write_row, column=3).value = row_list['USA']['total']
            dest_reach_net_sheet.cell(row=write_row, column=4).value = row_list['USA']['total_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['USA']['total_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=5).value = row_list['USA']['total_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=6).value = row_list['USA']['total_frequency_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=7).value = row_list['USA']['target']
            dest_reach_net_sheet.cell(row=write_row, column=8).value = row_list['USA']['target_pct']
            if equiv:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['USA']['target_impressions']
            else:
                dest_reach_net_sheet.cell(row=write_row, column=9).value = row_list['USA']['target_impressions_unequiv']
            dest_reach_net_sheet.cell(row=write_row, column=10).value = row_list['USA']['target_frequency_unequiv']
            write_row += 1

    summary_wb.save(filename)

    return True


def process_powerpoint_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_sheet = summary_wb.get_sheet_by_name('Summary Metrics')
    dest_pp_sheet = summary_wb.get_sheet_by_name('Powerpoint Data')
    source_rows = source_sheet.max_row

    for row_num in range(1, source_rows):
        if source_sheet.cell(row=row_num, column=1).value == 'Total':
            dest_pp_sheet.cell(row=4,column=3).value = source_sheet.cell(row=row_num, column=4).value
            if equiv:
                dest_pp_sheet.cell(row=5,column=3).value = source_sheet.cell(row=row_num, column=3).value
            else:
                dest_pp_sheet.cell(row=5,column=3).value = (source_sheet.cell(row=row_num, column=6).value * source_sheet.cell(row=row_num, column=3).value) / source_sheet.cell(row=row_num, column=5).value
            if equiv:
                dest_pp_sheet.cell(row=8,column=3).value = source_sheet.cell(row=row_num, column=17).value
            else:
                dest_pp_sheet.cell(row=8,column=3).value = source_sheet.cell(row=row_num, column=18).value
            dest_pp_sheet.cell(row=9,column=3).value = source_sheet.cell(row=row_num, column=32).value
            dest_pp_sheet.cell(row=10,column=3).value = source_sheet.cell(row=row_num, column=20).value
            if equiv:
                dest_pp_sheet.cell(row=11,column=3).value = source_sheet.cell(row=row_num, column=27).value
            else:
                dest_pp_sheet.cell(row=11,column=3).value = source_sheet.cell(row=row_num, column=28).value
            dest_pp_sheet.cell(row=12,column=3).value = source_sheet.cell(row=7, column=2).value
            dest_pp_sheet.cell(row=13,column=3).value = source_sheet.cell(row=15, column=2).value
            dest_pp_sheet.cell(row=14,column=3).value = source_sheet.cell(row=row_num, column=19).value
            dest_pp_sheet.cell(row=43,column=3).value = source_sheet.cell(row=row_num, column=29).value
            dest_pp_sheet.cell(row=68,column=3).value = source_sheet.cell(row=row_num, column=20).value
            break

    start_row = 0
    end_row = 0
    for row_num in range(2, source_rows):
        if source_sheet.cell(row=row_num, column=1).value == 'Network Summary':
            start_row = row_num

        if not source_sheet.cell(row=row_num, column=1).value and start_row:
            end_row = row_num + 1

    for row_num in range(start_row, end_row):
        if source_sheet.cell(row=row_num, column=1).value == 'NBC':
            if equiv:
                dest_pp_sheet.cell(row=25, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=51, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=25, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=51, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=76, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Bravo':
            if equiv:
                dest_pp_sheet.cell(row=18, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=44, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=18, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=44, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=69, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'CNBC':
            if equiv:
                dest_pp_sheet.cell(row=19, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=45, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=19, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=45, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=70, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Chiller':
            if equiv:
                dest_pp_sheet.cell(row=20, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=46, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=20, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=46, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=71, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'E!':
            if equiv:
                dest_pp_sheet.cell(row=21, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=47, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=21, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=47, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=72, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Esquire':
            if equiv:
                dest_pp_sheet.cell(row=22, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=48, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=22, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=48, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=73, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Golf Channel':
            if equiv:
                dest_pp_sheet.cell(row=23, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=49, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=23, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=49, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=74, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'MSNBC':
            if equiv:
                dest_pp_sheet.cell(row=24, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=50, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=24, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=50, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=75, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBCSN':
            if equiv:
                dest_pp_sheet.cell(row=26, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=52, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=26, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=52, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=77, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Oxygen':
            if equiv:
                dest_pp_sheet.cell(row=27, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=53, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=27, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=53, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=78, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'Syfy':
            if equiv:
                dest_pp_sheet.cell(row=28, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=54, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=28, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=54, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=79, column=3).value = source_sheet.cell(row=row_num, column=20).value
        if source_sheet.cell(row=row_num, column=1).value == 'USA':
            if equiv:
                dest_pp_sheet.cell(row=29, column=3).value = source_sheet.cell(row=row_num, column=17).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=55, column=3).value = source_sheet.cell(row=row_num, column=29).value
            else:
                dest_pp_sheet.cell(row=29, column=3).value = source_sheet.cell(row=row_num, column=18).value / dest_pp_sheet.cell(row=8,column=3).value
                dest_pp_sheet.cell(row=55, column=3).value = source_sheet.cell(row=row_num, column=30).value
            dest_pp_sheet.cell(row=80, column=3).value = source_sheet.cell(row=row_num, column=20).value

    source_sheet = summary_wb.get_sheet_by_name('Program Metrics')
    source_rows = source_sheet.max_row

    programs_obj = {}
    networks_obj = {}
    for row_num in range(2, source_rows + 1):
        networks_obj[source_sheet.cell(row=row_num,column=1).value] = 1
        programs_obj[source_sheet.cell(row=row_num,column=18).value] = 1

    dest_pp_sheet.cell(row=6,column=3).value = len(programs_obj.items())
    dest_pp_sheet.cell(row=7,column=3).value = len(networks_obj.items())

    source_sheet = summary_wb.get_sheet_by_name('Network Daypart')
    source_rows = source_sheet.max_row

    nbc_sum = 0
    nbc_morning = 0
    nbc_daytime = 0
    nbc_early = 0
    nbc_prime = 0
    nbc_late = 0
    nbc_overnight = 0
    for row_num in range(2, source_rows + 1):
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Morning':
            if equiv:
                nbc_morning = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=59, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_morning = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=59, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=84, column=3).value = source_sheet.cell(row=row_num, column=21).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Daytime':
            if equiv:
                nbc_daytime = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=60, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_daytime = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=60, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=85, column=3).value = source_sheet.cell(row=row_num, column=21).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Early Fringe':
            if equiv:
                nbc_early = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=61, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_early = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=61, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=86, column=3).value = source_sheet.cell(row=row_num, column=21).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Prime':
            if equiv:
                nbc_prime = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=62, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_prime = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=62, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=87, column=3).value = source_sheet.cell(row=row_num, column=21).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Late Night':
            if equiv:
                nbc_late = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=63, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_late = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=63, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=88, column=3).value = source_sheet.cell(row=row_num, column=21).value
        if source_sheet.cell(row=row_num, column=1).value == 'NBC' and source_sheet.cell(row=row_num, column=2).value == 'Overnight':
            if equiv:
                nbc_overnight = source_sheet.cell(row=row_num, column=18).value
                nbc_sum += source_sheet.cell(row=row_num, column=18).value
                dest_pp_sheet.cell(row=64, column=3).value = source_sheet.cell(row=row_num, column=30).value
            else:
                nbc_overnight = source_sheet.cell(row=row_num, column=19).value
                nbc_sum += source_sheet.cell(row=row_num, column=19).value
                dest_pp_sheet.cell(row=64, column=3).value = source_sheet.cell(row=row_num, column=31).value
            dest_pp_sheet.cell(row=89, column=3).value = source_sheet.cell(row=row_num, column=21).value

    dest_pp_sheet.cell(row=34, column=3).value = nbc_morning / nbc_sum if nbc_sum > 0 else 0
    dest_pp_sheet.cell(row=35, column=3).value = nbc_daytime / nbc_sum if nbc_sum > 0 else 0
    dest_pp_sheet.cell(row=36, column=3).value = nbc_early / nbc_sum if nbc_sum > 0 else 0
    dest_pp_sheet.cell(row=37, column=3).value = nbc_prime / nbc_sum if nbc_sum > 0 else 0
    dest_pp_sheet.cell(row=38, column=3).value = nbc_late / nbc_sum if nbc_sum > 0 else 0
    dest_pp_sheet.cell(row=39, column=3).value = nbc_overnight / nbc_sum if nbc_sum > 0 else 0

    summary_wb._active_sheet_index = 6

    summary_wb.save(filename)
    return True


def process_appendix_tab(filename, equiv):
    global source_wb
    global summary_wb

    source_sheet = summary_wb.get_sheet_by_name('Network Daypart')
    dest_pp_sheet = summary_wb.get_sheet_by_name('Appendix')
    source_rows = source_sheet.max_row
    day_net_obj = {'Morning':{}, 'Daytime':{}, 'Early Fringe':{}, 'Prime':{}, 'Late Night': {}, 'Overnight': {}}

    for row_num in range(2, source_rows + 1):
        if day_net_obj.has_key(source_sheet.cell(row=row_num, column=2).value):
            if equiv:
                day_net_obj[source_sheet.cell(row=row_num, column=2).value][source_sheet.cell(row=row_num, column=1).value] =\
                    {'target_impressions':source_sheet.cell(row=row_num, column=18).value,
                     'target_index':source_sheet.cell(row=row_num, column=30).value,
                     'target_reach':source_sheet.cell(row=row_num, column=20).value,
                     'target_frequency':source_sheet.cell(row=row_num, column=27).value,
                     'tCPM':source_sheet.cell(row=row_num, column=33).value}
            else:
                day_net_obj[source_sheet.cell(row=row_num, column=2).value][source_sheet.cell(row=row_num, column=1).value] =\
                    {'target_impressions':source_sheet.cell(row=row_num, column=19).value,
                     'target_index':source_sheet.cell(row=row_num, column=31).value,
                     'target_reach':source_sheet.cell(row=row_num, column=20).value,
                     'target_frequency':source_sheet.cell(row=row_num, column=27).value,
                     'tCPM':source_sheet.cell(row=row_num, column=33).value}
        else:
            if equiv:
                day_net_obj[source_sheet.cell(row=row_num, column=2).value] = {source_sheet.cell(row=row_num, column=1).value:
                                                                                   {'target_impressions':source_sheet.cell(row=row_num, column=18).value,
                                                                                    'target_index':source_sheet.cell(row=row_num, column=30).value,
                                                                                    'target_reach':source_sheet.cell(row=row_num, column=20).value,
                                                                                    'target_frequency':source_sheet.cell(row=row_num, column=27).value,
                                                                                    'tCPM':source_sheet.cell(row=row_num, column=33).value}}
            else:
                day_net_obj[source_sheet.cell(row=row_num, column=2).value] = {source_sheet.cell(row=row_num, column=1).value:
                                                                                   {'target_impressions':source_sheet.cell(row=row_num, column=19).value,
                                                                                    'target_index':source_sheet.cell(row=row_num, column=31).value,
                                                                                    'target_reach':source_sheet.cell(row=row_num, column=20).value,
                                                                                    'target_frequency':source_sheet.cell(row=row_num, column=27).value,
                                                                                    'tCPM':source_sheet.cell(row=row_num, column=33).value}}

    source_sheet = source_wb.get_sheet_by_name('Summary')
    source_rows = source_sheet.max_row
    net_obj = {}

    for row_num in range(2, source_rows + 1):
        if equiv:
            net_obj[source_sheet.cell(row=row_num, column=1).value] = {'target_impressions':source_sheet.cell(row=row_num, column=17).value,
                                                                       'target_index':source_sheet.cell(row=row_num, column=29).value,
                                                                       'target_reach':source_sheet.cell(row=row_num, column=19).value,
                                                                       'target_frequency':source_sheet.cell(row=row_num, column=26).value,
                                                                       'tCPM':source_sheet.cell(row=row_num, column=32).value}
        else:
            net_obj[source_sheet.cell(row=row_num, column=1).value] = {'target_impressions':source_sheet.cell(row=row_num, column=18).value,
                                                                       'target_index':source_sheet.cell(row=row_num, column=30).value,
                                                                       'target_reach':source_sheet.cell(row=row_num, column=19).value,
                                                                       'target_frequency':source_sheet.cell(row=row_num, column=26).value,
                                                                       'tCPM':source_sheet.cell(row=row_num, column=32).value}

    # Morning
    mo_im_total = 0
    mo_re_total = 0
    if day_net_obj['Morning'].has_key('Bravo'):
        dest_pp_sheet['B16'] = day_net_obj['Morning']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B39'] = day_net_obj['Morning']['Bravo']['target_index']
        dest_pp_sheet['B63'] = day_net_obj['Morning']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B88'] = day_net_obj['Morning']['Bravo']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('CNBC'):
        dest_pp_sheet['C16'] = day_net_obj['Morning']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C39'] = day_net_obj['Morning']['CNBC']['target_index']
        dest_pp_sheet['C63'] = day_net_obj['Morning']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C88'] = day_net_obj['Morning']['CNBC']['target_frequency']
        mo_im_total += day_net_obj['Morning']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('Chiller'):
        dest_pp_sheet['D16'] = day_net_obj['Morning']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D39'] = day_net_obj['Morning']['Chiller']['target_index']
        dest_pp_sheet['D63'] = day_net_obj['Morning']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D88'] = day_net_obj['Morning']['Chiller']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('E!'):
        dest_pp_sheet['E16'] = day_net_obj['Morning']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E39'] = day_net_obj['Morning']['E!']['target_index']
        dest_pp_sheet['E63'] = day_net_obj['Morning']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E88'] = day_net_obj['Morning']['E!']['target_frequency']
        mo_im_total += day_net_obj['Morning']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('Esquire'):
        dest_pp_sheet['F16'] = day_net_obj['Morning']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F39'] = day_net_obj['Morning']['Esquire']['target_index']
        dest_pp_sheet['F63'] = day_net_obj['Morning']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F88'] = day_net_obj['Morning']['Esquire']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('Golf Channel'):
        dest_pp_sheet['G16'] = day_net_obj['Morning']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G39'] = day_net_obj['Morning']['Golf Channel']['target_index']
        dest_pp_sheet['G63'] = day_net_obj['Morning']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G88'] = day_net_obj['Morning']['Golf Channel']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('NBC'):
        dest_pp_sheet['H16'] = day_net_obj['Morning']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H39'] = day_net_obj['Morning']['NBC']['target_index']
        dest_pp_sheet['H63'] = day_net_obj['Morning']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        mo_im_total += day_net_obj['Morning']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('NBCSN'):
        dest_pp_sheet['I16'] = day_net_obj['Morning']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I39'] = day_net_obj['Morning']['NBCSN']['target_index']
        dest_pp_sheet['I63'] = day_net_obj['Morning']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I88'] = day_net_obj['Morning']['NBCSN']['target_frequency']
        mo_im_total += day_net_obj['Morning']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('Oxygen'):
        dest_pp_sheet['J16'] = day_net_obj['Morning']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J39'] = day_net_obj['Morning']['Oxygen']['target_index']
        dest_pp_sheet['J63'] = day_net_obj['Morning']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J88'] = day_net_obj['Morning']['Oxygen']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('Syfy'):
        dest_pp_sheet['K16'] = day_net_obj['Morning']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K39'] = day_net_obj['Morning']['Syfy']['target_index']
        dest_pp_sheet['K63'] = day_net_obj['Morning']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K88'] = day_net_obj['Morning']['Syfy']['target_frequency']
        mo_im_total += day_net_obj['Morning']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('USA'):
        dest_pp_sheet['L16'] = day_net_obj['Morning']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L39'] = day_net_obj['Morning']['USA']['target_index']
        dest_pp_sheet['L63'] = day_net_obj['Morning']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L88'] = day_net_obj['Morning']['USA']['target_frequency']
        mo_im_total += day_net_obj['Morning']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Morning'].has_key('MSNBC'):
        dest_pp_sheet['M16'] = day_net_obj['Morning']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M39'] = day_net_obj['Morning']['MSNBC']['target_index']
        dest_pp_sheet['M63'] = day_net_obj['Morning']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M88'] = day_net_obj['Morning']['MSNBC']['target_frequency']
        mo_im_total += day_net_obj['Morning']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Morning']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']

    # Daytime
    dy_im_total = 0
    dy_re_total = 0
    if day_net_obj['Daytime'].has_key('Bravo'):
        dest_pp_sheet['B17'] = day_net_obj['Daytime']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B40'] = day_net_obj['Daytime']['Bravo']['target_index']
        dest_pp_sheet['B64'] = day_net_obj['Daytime']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B89'] = day_net_obj['Daytime']['Bravo']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('CNBC'):
        dest_pp_sheet['C17'] = day_net_obj['Daytime']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C40'] = day_net_obj['Daytime']['CNBC']['target_index']
        dest_pp_sheet['C64'] = day_net_obj['Daytime']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C89'] = day_net_obj['Daytime']['CNBC']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('Chiller'):
        dest_pp_sheet['D17'] = day_net_obj['Daytime']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D40'] = day_net_obj['Daytime']['Chiller']['target_index']
        dest_pp_sheet['D64'] = day_net_obj['Daytime']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D89'] = day_net_obj['Daytime']['Chiller']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('E!'):
        dest_pp_sheet['E17'] = day_net_obj['Daytime']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E40'] = day_net_obj['Daytime']['E!']['target_index']
        dest_pp_sheet['E64'] = day_net_obj['Daytime']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E89'] = day_net_obj['Daytime']['E!']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('Esquire'):
        dest_pp_sheet['F17'] = day_net_obj['Daytime']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F40'] = day_net_obj['Daytime']['Esquire']['target_index']
        dest_pp_sheet['F64'] = day_net_obj['Daytime']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F89'] = day_net_obj['Daytime']['Esquire']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('Golf Channel'):
        dest_pp_sheet['G17'] = day_net_obj['Daytime']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G40'] = day_net_obj['Daytime']['Golf Channel']['target_index']
        dest_pp_sheet['G64'] = day_net_obj['Daytime']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G89'] = day_net_obj['Daytime']['Golf Channel']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('NBC'):
        dest_pp_sheet['H17'] = day_net_obj['Daytime']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H40'] = day_net_obj['Daytime']['NBC']['target_index']
        dest_pp_sheet['H64'] = day_net_obj['Daytime']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['H89'] = day_net_obj['Daytime']['NBC']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('NBCSN'):
        dest_pp_sheet['I17'] = day_net_obj['Daytime']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I40'] = day_net_obj['Daytime']['NBCSN']['target_index']
        dest_pp_sheet['I64'] = day_net_obj['Daytime']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I89'] = day_net_obj['Daytime']['NBCSN']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('Oxygen'):
        dest_pp_sheet['J17'] = day_net_obj['Daytime']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J40'] = day_net_obj['Daytime']['Oxygen']['target_index']
        dest_pp_sheet['J64'] = day_net_obj['Daytime']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J89'] = day_net_obj['Daytime']['Oxygen']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('Syfy'):
        dest_pp_sheet['K17'] = day_net_obj['Daytime']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K40'] = day_net_obj['Daytime']['Syfy']['target_index']
        dest_pp_sheet['K64'] = day_net_obj['Daytime']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K89'] = day_net_obj['Daytime']['Syfy']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('USA'):
        dest_pp_sheet['L17'] = day_net_obj['Daytime']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L40'] = day_net_obj['Daytime']['USA']['target_index']
        dest_pp_sheet['L64'] = day_net_obj['Daytime']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L89'] = day_net_obj['Daytime']['USA']['target_frequency']
        dy_im_total += day_net_obj['Daytime']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dy_re_total += day_net_obj['Daytime']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Daytime'].has_key('MSNBC'):
        dest_pp_sheet['M17'] = day_net_obj['Daytime']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M40'] = day_net_obj['Daytime']['MSNBC']['target_index']
        dest_pp_sheet['M64'] = day_net_obj['Daytime']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M89'] = day_net_obj['Daytime']['MSNBC']['target_frequency']
        mo_im_total += day_net_obj['Daytime']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        mo_re_total += day_net_obj['Daytime']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']

    # Early Fringe
    fe_im_total = 0
    fe_re_total = 0
    if day_net_obj['Early Fringe'].has_key('Bravo'):
        dest_pp_sheet['B18'] = day_net_obj['Early Fringe']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B41'] = day_net_obj['Early Fringe']['Bravo']['target_index']
        dest_pp_sheet['B65'] = day_net_obj['Early Fringe']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B90'] = day_net_obj['Early Fringe']['Bravo']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('CNBC'):
        dest_pp_sheet['C18'] = day_net_obj['Early Fringe']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C41'] = day_net_obj['Early Fringe']['CNBC']['target_index']
        dest_pp_sheet['C65'] = day_net_obj['Early Fringe']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C90'] = day_net_obj['Early Fringe']['CNBC']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('Chiller'):
        dest_pp_sheet['D18'] = day_net_obj['Early Fringe']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D41'] = day_net_obj['Early Fringe']['Chiller']['target_index']
        dest_pp_sheet['D65'] = day_net_obj['Early Fringe']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D90'] = day_net_obj['Early Fringe']['Chiller']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('E!'):
        dest_pp_sheet['E18'] = day_net_obj['Early Fringe']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E41'] = day_net_obj['Early Fringe']['E!']['target_index']
        dest_pp_sheet['E65'] = day_net_obj['Early Fringe']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E90'] = day_net_obj['Early Fringe']['E!']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('Esquire'):
        dest_pp_sheet['F18'] = day_net_obj['Early Fringe']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F41'] = day_net_obj['Early Fringe']['Esquire']['target_index']
        dest_pp_sheet['F65'] = day_net_obj['Early Fringe']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F90'] = day_net_obj['Early Fringe']['Esquire']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('Golf Channel'):
        dest_pp_sheet['G18'] = day_net_obj['Early Fringe']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G41'] = day_net_obj['Early Fringe']['Golf Channel']['target_index']
        dest_pp_sheet['G65'] = day_net_obj['Early Fringe']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G90'] = day_net_obj['Early Fringe']['Golf Channel']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('NBC'):
        dest_pp_sheet['H18'] = day_net_obj['Early Fringe']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H41'] = day_net_obj['Early Fringe']['NBC']['target_index']
        dest_pp_sheet['H65'] = day_net_obj['Early Fringe']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['H90'] = day_net_obj['Early Fringe']['NBC']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('NBCSN'):
        dest_pp_sheet['I18'] = day_net_obj['Early Fringe']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I41'] = day_net_obj['Early Fringe']['NBCSN']['target_index']
        dest_pp_sheet['I65'] = day_net_obj['Early Fringe']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I90'] = day_net_obj['Early Fringe']['NBCSN']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('Oxygen'):
        dest_pp_sheet['J18'] = day_net_obj['Early Fringe']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J41'] = day_net_obj['Early Fringe']['Oxygen']['target_index']
        dest_pp_sheet['J65'] = day_net_obj['Early Fringe']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J90'] = day_net_obj['Early Fringe']['Oxygen']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('Syfy'):
        dest_pp_sheet['K18'] = day_net_obj['Early Fringe']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K41'] = day_net_obj['Early Fringe']['Syfy']['target_index']
        dest_pp_sheet['K65'] = day_net_obj['Early Fringe']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K90'] = day_net_obj['Early Fringe']['Syfy']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('USA'):
        dest_pp_sheet['L18'] = day_net_obj['Early Fringe']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L41'] = day_net_obj['Early Fringe']['USA']['target_index']
        dest_pp_sheet['L65'] = day_net_obj['Early Fringe']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L90'] = day_net_obj['Early Fringe']['USA']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Early Fringe'].has_key('MSNBC'):
        dest_pp_sheet['M18'] = day_net_obj['Early Fringe']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M41'] = day_net_obj['Early Fringe']['MSNBC']['target_index']
        dest_pp_sheet['M65'] = day_net_obj['Early Fringe']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M90'] = day_net_obj['Early Fringe']['MSNBC']['target_frequency']
        fe_im_total += day_net_obj['Early Fringe']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        fe_re_total += day_net_obj['Early Fringe']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']



    # Prime
    pr_im_total = 0
    pr_re_total = 0
    if day_net_obj['Prime'].has_key('Bravo'):
        dest_pp_sheet['B19'] = day_net_obj['Prime']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B42'] = day_net_obj['Prime']['Bravo']['target_index']
        dest_pp_sheet['B66'] = day_net_obj['Prime']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B91'] = day_net_obj['Prime']['Bravo']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('CNBC'):
        dest_pp_sheet['C19'] = day_net_obj['Prime']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C42'] = day_net_obj['Prime']['CNBC']['target_index']
        dest_pp_sheet['C66'] = day_net_obj['Prime']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C91'] = day_net_obj['Prime']['CNBC']['target_frequency']
        pr_im_total += day_net_obj['Prime']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('Chiller'):
        dest_pp_sheet['D19'] = day_net_obj['Prime']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D42'] = day_net_obj['Prime']['Chiller']['target_index']
        dest_pp_sheet['D66'] = day_net_obj['Prime']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D91'] = day_net_obj['Prime']['Chiller']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('E!'):
        dest_pp_sheet['E19'] = day_net_obj['Prime']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E42'] = day_net_obj['Prime']['E!']['target_index']
        dest_pp_sheet['E66'] = day_net_obj['Prime']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E91'] = day_net_obj['Prime']['E!']['target_frequency']
        pr_im_total += day_net_obj['Prime']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('Esquire'):
        dest_pp_sheet['F19'] = day_net_obj['Prime']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F42'] = day_net_obj['Prime']['Esquire']['target_index']
        dest_pp_sheet['F66'] = day_net_obj['Prime']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F91'] = day_net_obj['Prime']['Esquire']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('Golf Channel'):
        dest_pp_sheet['G19'] = day_net_obj['Prime']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G42'] = day_net_obj['Prime']['Golf Channel']['target_index']
        dest_pp_sheet['G66'] = day_net_obj['Prime']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G91'] = day_net_obj['Prime']['Golf Channel']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('NBC'):
        dest_pp_sheet['H19'] = day_net_obj['Prime']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H42'] = day_net_obj['Prime']['NBC']['target_index']
        dest_pp_sheet['H66'] = day_net_obj['Prime']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['H91'] = day_net_obj['Prime']['NBC']['target_frequency']
        pr_im_total += day_net_obj['Prime']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('NBCSN'):
        dest_pp_sheet['I19'] = day_net_obj['Prime']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I42'] = day_net_obj['Prime']['NBCSN']['target_index']
        dest_pp_sheet['I66'] = day_net_obj['Prime']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I91'] = day_net_obj['Prime']['NBCSN']['target_frequency']
        pr_im_total += day_net_obj['Prime']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('Oxygen'):
        dest_pp_sheet['J19'] = day_net_obj['Prime']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J42'] = day_net_obj['Prime']['Oxygen']['target_index']
        dest_pp_sheet['J66'] = day_net_obj['Prime']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J91'] = day_net_obj['Prime']['Oxygen']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('Syfy'):
        dest_pp_sheet['K19'] = day_net_obj['Prime']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K42'] = day_net_obj['Prime']['Syfy']['target_index']
        dest_pp_sheet['K66'] = day_net_obj['Prime']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K91'] = day_net_obj['Prime']['Syfy']['target_frequency']
        pr_im_total += day_net_obj['Prime']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('USA'):
        dest_pp_sheet['L19'] = day_net_obj['Prime']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L42'] = day_net_obj['Prime']['USA']['target_index']
        dest_pp_sheet['L66'] = day_net_obj['Prime']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L91'] = day_net_obj['Prime']['USA']['target_frequency']
        pr_im_total += day_net_obj['Prime']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Prime'].has_key('MSNBC'):
        dest_pp_sheet['M19'] = day_net_obj['Prime']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M42'] = day_net_obj['Prime']['MSNBC']['target_index']
        dest_pp_sheet['M66'] = day_net_obj['Prime']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M91'] = day_net_obj['Prime']['MSNBC']['target_frequency']
        pr_im_total += day_net_obj['Prime']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        pr_re_total += day_net_obj['Prime']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']

    # Late Night
    ln_im_total = 0
    ln_re_total = 0
    if day_net_obj['Late Night'].has_key('Bravo'):
        dest_pp_sheet['B20'] = day_net_obj['Late Night']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B43'] = day_net_obj['Late Night']['Bravo']['target_index']
        dest_pp_sheet['B67'] = day_net_obj['Late Night']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B92'] = day_net_obj['Late Night']['Bravo']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('CNBC'):
        dest_pp_sheet['C20'] = day_net_obj['Late Night']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C43'] = day_net_obj['Late Night']['CNBC']['target_index']
        dest_pp_sheet['C67'] = day_net_obj['Late Night']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C92'] = day_net_obj['Late Night']['CNBC']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('Chiller'):
        dest_pp_sheet['D20'] = day_net_obj['Late Night']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D43'] = day_net_obj['Late Night']['Chiller']['target_index']
        dest_pp_sheet['D67'] = day_net_obj['Late Night']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D92'] = day_net_obj['Late Night']['Chiller']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('E!'):
        dest_pp_sheet['E20'] = day_net_obj['Late Night']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E43'] = day_net_obj['Late Night']['E!']['target_index']
        dest_pp_sheet['E67'] = day_net_obj['Late Night']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E92'] = day_net_obj['Late Night']['E!']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('Esquire'):
        dest_pp_sheet['F20'] = day_net_obj['Late Night']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F43'] = day_net_obj['Late Night']['Esquire']['target_index']
        dest_pp_sheet['F67'] = day_net_obj['Late Night']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F92'] = day_net_obj['Late Night']['Esquire']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('Golf Channel'):
        dest_pp_sheet['G20'] = day_net_obj['Late Night']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G43'] = day_net_obj['Late Night']['Golf Channel']['target_index']
        dest_pp_sheet['G67'] = day_net_obj['Late Night']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G92'] = day_net_obj['Late Night']['Golf Channel']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('NBC'):
        dest_pp_sheet['H20'] = day_net_obj['Late Night']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H43'] = day_net_obj['Late Night']['NBC']['target_index']
        dest_pp_sheet['H67'] = day_net_obj['Late Night']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['H92'] = day_net_obj['Late Night']['NBC']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('NBCSN'):
        dest_pp_sheet['I20'] = day_net_obj['Late Night']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I43'] = day_net_obj['Late Night']['NBCSN']['target_index']
        dest_pp_sheet['I67'] = day_net_obj['Late Night']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I92'] = day_net_obj['Late Night']['NBCSN']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('Oxygen'):
        dest_pp_sheet['J20'] = day_net_obj['Late Night']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J43'] = day_net_obj['Late Night']['Oxygen']['target_index']
        dest_pp_sheet['J67'] = day_net_obj['Late Night']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J92'] = day_net_obj['Late Night']['Oxygen']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('Syfy'):
        dest_pp_sheet['K20'] = day_net_obj['Late Night']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K43'] = day_net_obj['Late Night']['Syfy']['target_index']
        dest_pp_sheet['K67'] = day_net_obj['Late Night']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K92'] = day_net_obj['Late Night']['Syfy']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('USA'):
        dest_pp_sheet['L20'] = day_net_obj['Late Night']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L43'] = day_net_obj['Late Night']['USA']['target_index']
        dest_pp_sheet['L67'] = day_net_obj['Late Night']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L92'] = day_net_obj['Late Night']['USA']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj['Late Night'].has_key('MSNBC'):
        dest_pp_sheet['M20'] = day_net_obj['Late Night']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M43'] = day_net_obj['Late Night']['MSNBC']['target_index']
        dest_pp_sheet['M67'] = day_net_obj['Late Night']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M92'] = day_net_obj['Late Night']['MSNBC']['target_frequency']
        ln_im_total += day_net_obj['Late Night']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        ln_re_total += day_net_obj['Late Night']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']

    # Overnight
    on_im_total = 0
    on_re_total = 0
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Bravo'):
        dest_pp_sheet['B21'] = day_net_obj['Overnight']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['B44'] = day_net_obj['Overnight']['Bravo']['target_index']
        dest_pp_sheet['B68'] = day_net_obj['Overnight']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['B93'] = day_net_obj['Overnight']['Bravo']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Bravo']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Bravo']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('CNBC'):
        dest_pp_sheet['C21'] = day_net_obj['Overnight']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['C44'] = day_net_obj['Overnight']['CNBC']['target_index']
        dest_pp_sheet['C68'] = day_net_obj['Overnight']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['C93'] = day_net_obj['Overnight']['CNBC']['target_frequency']
        on_im_total += day_net_obj['Overnight']['CNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['CNBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Chiller'):
        dest_pp_sheet['D21'] = day_net_obj['Overnight']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['D44'] = day_net_obj['Overnight']['Chiller']['target_index']
        dest_pp_sheet['D68'] = day_net_obj['Overnight']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['D93'] = day_net_obj['Overnight']['Chiller']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Chiller']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Chiller']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('E!'):
        dest_pp_sheet['E21'] = day_net_obj['Overnight']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['E44'] = day_net_obj['Overnight']['E!']['target_index']
        dest_pp_sheet['E68'] = day_net_obj['Overnight']['E!']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['E93'] = day_net_obj['Overnight']['E!']['target_frequency']
        on_im_total += day_net_obj['Overnight']['E!']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['E!']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Esquire'):
        dest_pp_sheet['F21'] = day_net_obj['Overnight']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['F44'] = day_net_obj['Overnight']['Esquire']['target_index']
        dest_pp_sheet['F68'] = day_net_obj['Overnight']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['F93'] = day_net_obj['Overnight']['Esquire']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Esquire']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Esquire']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Golf Channel'):
        dest_pp_sheet['G21'] = day_net_obj['Overnight']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['G44'] = day_net_obj['Overnight']['Golf Channel']['target_index']
        dest_pp_sheet['G68'] = day_net_obj['Overnight']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['G93'] = day_net_obj['Overnight']['Golf Channel']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Golf Channel']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBC'):
        dest_pp_sheet['H21'] = day_net_obj['Overnight']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['H44'] = day_net_obj['Overnight']['NBC']['target_index']
        dest_pp_sheet['H68'] = day_net_obj['Overnight']['NBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['H93'] = day_net_obj['Overnight']['NBC']['target_frequency']
        on_im_total += day_net_obj['Overnight']['NBC']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['NBC']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBCSN'):
        dest_pp_sheet['I21'] = day_net_obj['Overnight']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['I44'] = day_net_obj['Overnight']['NBCSN']['target_index']
        dest_pp_sheet['I68'] = day_net_obj['Overnight']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['I93'] = day_net_obj['Overnight']['NBCSN']['target_frequency']
        on_im_total += day_net_obj['Overnight']['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['NBCSN']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Oxygen'):
        dest_pp_sheet['J21'] = day_net_obj['Overnight']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['J44'] = day_net_obj['Overnight']['Oxygen']['target_index']
        dest_pp_sheet['J68'] = day_net_obj['Overnight']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['J93'] = day_net_obj['Overnight']['Oxygen']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Oxygen']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('Syfy'):
        dest_pp_sheet['K21'] = day_net_obj['Overnight']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['K44'] = day_net_obj['Overnight']['Syfy']['target_index']
        dest_pp_sheet['K68'] = day_net_obj['Overnight']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['K93'] = day_net_obj['Overnight']['Syfy']['target_frequency']
        on_im_total += day_net_obj['Overnight']['Syfy']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['Syfy']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('USA'):
        dest_pp_sheet['L21'] = day_net_obj['Overnight']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['L44'] = day_net_obj['Overnight']['USA']['target_index']
        dest_pp_sheet['L68'] = day_net_obj['Overnight']['USA']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['L93'] = day_net_obj['Overnight']['USA']['target_frequency']
        on_im_total += day_net_obj['Overnight']['USA']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['USA']['target_reach'] / net_obj['Total']['target_reach']
    if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('MSNBC'):
        dest_pp_sheet['M21'] = day_net_obj['Overnight']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        dest_pp_sheet['M44'] = day_net_obj['Overnight']['MSNBC']['target_index']
        dest_pp_sheet['M68'] = day_net_obj['Overnight']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']
        dest_pp_sheet['M93'] = day_net_obj['Overnight']['MSNBC']['target_frequency']
        on_im_total += day_net_obj['Overnight']['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions']
        on_re_total += day_net_obj['Overnight']['MSNBC']['target_reach'] / net_obj['Total']['target_reach']

    dest_pp_sheet['N16'] = mo_im_total
    dest_pp_sheet['N17'] = dy_im_total
    dest_pp_sheet['N18'] = fe_im_total
    dest_pp_sheet['N19'] = pr_im_total
    dest_pp_sheet['N20'] = ln_im_total
    dest_pp_sheet['N21'] = on_im_total

    dest_pp_sheet['N63'] = mo_re_total
    dest_pp_sheet['N64'] = dy_re_total
    dest_pp_sheet['N65'] = fe_re_total
    dest_pp_sheet['N66'] = pr_re_total
    dest_pp_sheet['N67'] = ln_re_total
    dest_pp_sheet['N68'] = on_re_total

    dest_pp_sheet['N39'] = day_net_obj['Morning']['Total']['target_index'] if day_net_obj['Morning'].has_key('Total') else ''
    dest_pp_sheet['N40'] = day_net_obj['Daytime']['Total']['target_index'] if day_net_obj['Daytime'].has_key('Total') else ''
    dest_pp_sheet['N41'] = day_net_obj['Early Fringe']['Total']['target_index'] if day_net_obj['Early Fringe'].has_key('Total') else ''
    dest_pp_sheet['N42'] = day_net_obj['Prime']['Total']['target_index'] if day_net_obj['Prime'].has_key('Total') else ''
    dest_pp_sheet['N43'] = day_net_obj['Late Night']['Total']['target_index'] if day_net_obj['Late Night'].has_key('Total') else ''
    dest_pp_sheet['N44'] = day_net_obj['Overnight']['Total']['target_index'] if day_net_obj['Overnight'].has_key('Total') else ''

    dest_pp_sheet['N88'] = day_net_obj['Morning']['Total']['target_frequency'] if day_net_obj['Morning'].has_key('Total') else ''
    dest_pp_sheet['N89'] = day_net_obj['Daytime']['Total']['target_frequency'] if day_net_obj['Daytime'].has_key('Total') else ''
    dest_pp_sheet['N90'] = day_net_obj['Early Fringe']['Total']['target_frequency'] if day_net_obj['Early Fringe'].has_key('Total') else ''
    dest_pp_sheet['N91'] = day_net_obj['Prime']['Total']['target_frequency'] if day_net_obj['Prime'].has_key('Total') else ''
    dest_pp_sheet['N92'] = day_net_obj['Late Night']['Total']['target_frequency'] if day_net_obj['Late Night'].has_key('Total') else ''
    dest_pp_sheet['N93'] = day_net_obj['Overnight']['Total']['target_frequency'] if day_net_obj['Overnight'].has_key('Total') else ''

    # Totals
    dest_pp_sheet['B22'] = net_obj['Bravo']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C22'] = net_obj['CNBC']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['D22'] = net_obj['Chiller']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['E22'] = net_obj['E!']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['F22'] = net_obj['Esquire']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['G22'] = net_obj['Golf Channel']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['H22'] = net_obj['NBC']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('NBC') else ''
    dest_pp_sheet['I22'] = net_obj['NBCSN']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['J22'] = net_obj['Oxygen']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['K22'] = net_obj['Syfy']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['L22'] = net_obj['USA']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['M22'] = net_obj['MSNBC']['target_impressions'] / net_obj['Total']['target_impressions'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['N22'] = net_obj['Total']['target_impressions'] / net_obj['Total']['target_impressions']

    dest_pp_sheet['B45'] = net_obj['Bravo']['target_index'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C45'] = net_obj['CNBC']['target_index'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['D45'] = net_obj['Chiller']['target_index'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['E45'] = net_obj['E!']['target_index'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['F45'] = net_obj['Esquire']['target_index'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['G45'] = net_obj['Golf Channel']['target_index'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['H45'] = net_obj['NBC']['target_index'] if net_obj.has_key('NBC') else ''
    dest_pp_sheet['I45'] = net_obj['NBCSN']['target_index'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['J45'] = net_obj['Oxygen']['target_index'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['K45'] = net_obj['Syfy']['target_index'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['L45'] = net_obj['USA']['target_index'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['M45'] = net_obj['MSNBC']['target_index'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['N45'] = net_obj['Total']['target_index']

    dest_pp_sheet['B69'] = net_obj['Bravo']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C69'] = net_obj['CNBC']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['D69'] = net_obj['Chiller']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['E69'] = net_obj['E!']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['F69'] = net_obj['Esquire']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['G69'] = net_obj['Golf Channel']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['H69'] = net_obj['NBC']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('NBC') else ''
    dest_pp_sheet['I69'] = net_obj['NBCSN']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['J69'] = net_obj['Oxygen']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['K69'] = net_obj['Syfy']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['L69'] = net_obj['USA']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['M69'] = net_obj['MSNBC']['target_reach'] / net_obj['Total']['target_reach'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['N69'] = net_obj['Total']['target_reach'] / net_obj['Total']['target_reach']

    dest_pp_sheet['B94'] = net_obj['Bravo']['target_frequency'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C94'] = net_obj['CNBC']['target_frequency'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['D94'] = net_obj['Chiller']['target_frequency'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['E94'] = net_obj['E!']['target_frequency'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['F94'] = net_obj['Esquire']['target_frequency'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['G94'] = net_obj['Golf Channel']['target_frequency'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['H94'] = net_obj['NBC']['target_frequency'] if net_obj.has_key('NBC') else ''
    dest_pp_sheet['I94'] = net_obj['NBCSN']['target_frequency'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['J94'] = net_obj['Oxygen']['target_frequency'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['K94'] = net_obj['Syfy']['target_frequency'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['L94'] = net_obj['USA']['target_frequency'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['M94'] = net_obj['MSNBC']['target_frequency'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['N94'] = net_obj['Total']['target_frequency']

    dest_pp_sheet['C100'] = net_obj['Bravo']['target_frequency'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C101'] = net_obj['CNBC']['target_frequency'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['C102'] = net_obj['Chiller']['target_frequency'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['C103'] = net_obj['E!']['target_frequency'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['C104'] = net_obj['Esquire']['target_frequency'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['C105'] = net_obj['Golf Channel']['target_frequency'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['C106'] = net_obj['NBC']['target_frequency'] if net_obj.has_key('NBC') else ''
    dest_pp_sheet['C107'] = net_obj['NBCSN']['target_frequency'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['C108'] = net_obj['Oxygen']['target_frequency'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['C109'] = net_obj['Syfy']['target_frequency'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['C110'] = net_obj['USA']['target_frequency'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['C111'] = net_obj['MSNBC']['target_frequency'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['C99'] = net_obj['Total']['target_frequency']

    dest_pp_sheet['C115'] = day_net_obj['Morning']['NBC']['target_frequency'] if day_net_obj['Morning'].has_key('NBC') else ''
    dest_pp_sheet['C116'] = day_net_obj['Daytime']['NBC']['target_frequency'] if day_net_obj['Daytime'].has_key('NBC') else ''
    dest_pp_sheet['C117'] = day_net_obj['Early Fringe']['NBC']['target_frequency'] if day_net_obj['Early Fringe'].has_key('NBC') else ''
    dest_pp_sheet['C118'] = day_net_obj['Prime']['NBC']['target_frequency'] if day_net_obj['Prime'].has_key('NBC') else ''
    dest_pp_sheet['C119'] = day_net_obj['Late Night']['NBC']['target_frequency'] if day_net_obj['Late Night'].has_key('NBC') else ''
    dest_pp_sheet['C120'] = day_net_obj['Overnight']['NBC']['target_frequency'] if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBC') else ''

    dest_pp_sheet['A124'] = net_obj['Bravo']['tCPM'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['B124'] = net_obj['Bravo']['target_reach'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['C124'] = net_obj['Bravo']['target_frequency'] if net_obj.has_key('Bravo') else ''
    dest_pp_sheet['A125'] = net_obj['CNBC']['tCPM'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['B125'] = net_obj['CNBC']['target_reach'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['C125'] = net_obj['CNBC']['target_frequency'] if net_obj.has_key('CNBC') else ''
    dest_pp_sheet['A126'] = net_obj['Chiller']['tCPM'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['B126'] = net_obj['Chiller']['target_reach'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['C126'] = net_obj['Chiller']['target_frequency'] if net_obj.has_key('Chiller') else ''
    dest_pp_sheet['A127'] = net_obj['E!']['tCPM'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['B127'] = net_obj['E!']['target_reach'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['C127'] = net_obj['E!']['target_frequency'] if net_obj.has_key('E!') else ''
    dest_pp_sheet['A128'] = net_obj['Esquire']['tCPM'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['B128'] = net_obj['Esquire']['target_reach'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['C128'] = net_obj['Esquire']['target_frequency'] if net_obj.has_key('Esquire') else ''
    dest_pp_sheet['A129'] = net_obj['Golf Channel']['tCPM'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['B129'] = net_obj['Golf Channel']['target_reach'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['C129'] = net_obj['Golf Channel']['target_frequency'] if net_obj.has_key('Golf Channel') else ''
    dest_pp_sheet['A130'] = net_obj['NBCSN']['tCPM'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['B130'] = net_obj['NBCSN']['target_reach'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['C130'] = net_obj['NBCSN']['target_frequency'] if net_obj.has_key('NBCSN') else ''
    dest_pp_sheet['A131'] = net_obj['Oxygen']['tCPM'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['B131'] = net_obj['Oxygen']['target_reach'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['C131'] = net_obj['Oxygen']['target_frequency'] if net_obj.has_key('Oxygen') else ''
    dest_pp_sheet['A132'] = net_obj['Syfy']['tCPM'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['B132'] = net_obj['Syfy']['target_reach'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['C132'] = net_obj['Syfy']['target_frequency'] if net_obj.has_key('Syfy') else ''
    dest_pp_sheet['A133'] = net_obj['USA']['tCPM'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['B133'] = net_obj['USA']['target_reach'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['C133'] = net_obj['USA']['target_frequency'] if net_obj.has_key('USA') else ''
    dest_pp_sheet['A134'] = net_obj['MSNBC']['tCPM'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['B134'] = net_obj['MSNBC']['target_reach'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['C134'] = net_obj['MSNBC']['target_frequency'] if net_obj.has_key('MSNBC') else ''
    dest_pp_sheet['A135'] = day_net_obj['Morning']['NBC']['tCPM'] if day_net_obj['Morning'].has_key('NBC') else ''
    dest_pp_sheet['B135'] = day_net_obj['Morning']['NBC']['target_reach'] if day_net_obj['Morning'].has_key('NBC') else ''
    dest_pp_sheet['C135'] = day_net_obj['Morning']['NBC']['target_frequency'] if day_net_obj['Morning'].has_key('NBC') else ''
    dest_pp_sheet['A136'] = day_net_obj['Daytime']['NBC']['tCPM'] if day_net_obj['Daytime'].has_key('NBC') else ''
    dest_pp_sheet['B136'] = day_net_obj['Daytime']['NBC']['target_reach'] if day_net_obj['Daytime'].has_key('NBC') else ''
    dest_pp_sheet['C136'] = day_net_obj['Daytime']['NBC']['target_frequency'] if day_net_obj['Daytime'].has_key('NBC') else ''
    dest_pp_sheet['A137'] = day_net_obj['Early Fringe']['NBC']['tCPM'] if day_net_obj['Early Fringe'].has_key('NBC') else ''
    dest_pp_sheet['B137'] = day_net_obj['Early Fringe']['NBC']['target_reach'] if day_net_obj['Early Fringe'].has_key('NBC') else ''
    dest_pp_sheet['C137'] = day_net_obj['Early Fringe']['NBC']['target_frequency'] if day_net_obj['Early Fringe'].has_key('NBC') else ''
    dest_pp_sheet['A138'] = day_net_obj['Prime']['NBC']['tCPM'] if day_net_obj['Prime'].has_key('NBC') else ''
    dest_pp_sheet['B138'] = day_net_obj['Prime']['NBC']['target_reach'] if day_net_obj['Prime'].has_key('NBC') else ''
    dest_pp_sheet['C138'] = day_net_obj['Prime']['NBC']['target_frequency'] if day_net_obj['Prime'].has_key('NBC') else ''
    dest_pp_sheet['A139'] = day_net_obj['Late Night']['NBC']['tCPM'] if day_net_obj.has_key('Overnight') and day_net_obj['Late Night'].has_key('NBC') else ''
    dest_pp_sheet['B139'] = day_net_obj['Late Night']['NBC']['target_reach'] if day_net_obj.has_key('Overnight') and day_net_obj['Late Night'].has_key('NBC') else ''
    dest_pp_sheet['C139'] = day_net_obj['Late Night']['NBC']['target_frequency'] if day_net_obj.has_key('Overnight') and day_net_obj['Late Night'].has_key('NBC') else ''
    dest_pp_sheet['A140'] = day_net_obj['Overnight']['NBC']['tCPM'] if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBC') else ''
    dest_pp_sheet['B140'] = day_net_obj['Overnight']['NBC']['target_reach'] if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBC') else ''
    dest_pp_sheet['C140'] = day_net_obj['Overnight']['NBC']['target_frequency'] if day_net_obj.has_key('Overnight') and day_net_obj['Overnight'].has_key('NBC') else ''

    summary_wb.save(filename)
    return True


def move_when_done(processed, summary_equiv, summary_unequiv):
    shutil.move('./preprocessed/'+processed, './processed/'+processed)
    shutil.move(summary_equiv, './summaries/'+summary_equiv)
    shutil.move(summary_unequiv, './summaries/'+summary_unequiv)
    return True

listing = glob.glob('./preprocessed/*')
for filename in listing:
    filename = os.path.basename(filename)
    if os.path.isfile('./preprocessed/' + filename):
        print "processing unequiv"
        new_filename_unequiv = setup(filename, False)
        if not new_filename_unequiv:
            print "error setting up " + filename
        print process_summary_tab(new_filename_unequiv, False)
        print process_Network_Daypart_tab(new_filename_unequiv, False)
        print process_frequency_distribution_tab(new_filename_unequiv, False)
        print process_reach_by_week_tab(new_filename_unequiv, False)
        print process_frequency_distribution_by_net_tab(new_filename_unequiv, False)
        print process_network_reach_tab(new_filename_unequiv, False)
        print process_powerpoint_tab(new_filename_unequiv, False)
        print process_appendix_tab(new_filename_unequiv, False)
        print "processing equiv"
        new_filename_equiv = setup(filename, True)
        if not new_filename_equiv:
            print "error setting up " + filename
        print process_summary_tab(new_filename_equiv, True)
        print process_Network_Daypart_tab(new_filename_equiv, True)
        print process_frequency_distribution_tab(new_filename_equiv, True)
        print process_reach_by_week_tab(new_filename_equiv, True)
        print process_frequency_distribution_by_net_tab(new_filename_equiv, True)
        print process_network_reach_tab(new_filename_equiv, True)
        print process_powerpoint_tab(new_filename_equiv, True)
        print process_appendix_tab(new_filename_equiv, True)
        print move_when_done(filename, new_filename_equiv, new_filename_unequiv)
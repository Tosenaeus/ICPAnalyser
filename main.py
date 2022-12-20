import regex as re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np
from scikit_posthocs import outliers_grubbs as grubbs
from scipy import stats
import math

rx = re.compile(r'"[^"]*"(*SKIP)(*FAIL)|,\s*')

print('Automatische Auswertung für Ergebnisse aus Varian ICP Expert')
print('Author: Hendrik Marx | Version: 1.0.1')
print('.csv im selben Ordner platzieren und Dateinamen ohne Endung eingeben')
file = input()

workbook = xlsxwriter.Workbook(file + '.xlsx')
overview = workbook.add_worksheet('Overview')
calibrate = workbook.add_worksheet('Calibration')
result = workbook.add_worksheet('Results')

bold = workbook.add_format({'bold': True})
red = workbook.add_format({'font_color': 'red'})

smpl = []  # lines concerning samples
clbrtn = []  # lines concerning calibration
lns = []  # elemental lines
smplnms = []  # sample names
ar = []  # numbers of Argon lines
arint = {}  # intensity of argon blanks
cal = []  # calibration steps
dlt = []  # samples to delete
dltnms = []  # names of the samples to be deleted
sel = []  # selection of samples
blk = []  # intensities of a single blank
blk2 = []  # list of stds of blk
blksd = {}  # blank standard deviation

# read and split into calibration and sample
with open(file + '.csv', 'r') as file:
    for line in file:
        if ':\\' in line:
            continue
        strp = line.strip()
        splt = rx.split(strp)

        if '"Solution Label"' in splt:
            sltnlbl = splt.index('"Solution Label"')
            tp = splt.index('"Type"')
            lmnt = splt.index('"Element"')
            flgs = splt.index('"Flags"')
            cnctrn = splt.index('"Soln Conc"')
            nt = splt.index('"Int"')
            dt = splt.index('"Date"')
            tm = splt.index('"Time"')
            nmrps = splt.index('"NumReps"')
        if splt[tp] == '"Bld"' or splt[tp] == '"Kal"':
            clbrtn.append(splt)
            if splt[tp] == '"Bld"' and 'Ar' in splt[lmnt]:
                arint[splt[lmnt]] = splt[nt]
        if splt[tp] == '"Probe"':
            smpl.append(splt)

# get measured elemental lines
for i in clbrtn:
    if lns == []:
        lns.append(i[lmnt])
    elif lns[0] == i[lmnt]:
        break
    else:
        lns.append(i[lmnt])

# get sample names
amount = int(len(smpl) / len(lns))
for i in range(0, amount):
    smplnms.append(smpl[i * len(lns)][sltnlbl])

# select samples
print(str(len(smplnms)) + ' Proben erkannt:')
print('Index: Probenname')
co = 0
for i in smplnms:
    print(str(co) + ': ' + i)
    co += 1
print(
    'Um Proben zu löschen und nicht zu berechnen, Index der zu löschenden Proben mit Komma getrennt und ohne Leerzeichen eingeben.' \
    + 'Um nichts zu löschen, ohne Eingabe Enter drücken.')
sel = input().split(',')
if sel != ['']:
    for i in sel:
        dlt.append(int(i))
    for i in dlt:
        dltnms.append(smplnms[i])
    co = 0
    nrdlt = len(smpl)
    while co < nrdlt:
        if smpl[co][sltnlbl] in dltnms:
            smpl.pop(co)
        else:
            co += 1
        nrdlt = len(smpl)

    # update sample names
    smplnms = []
    amount = int(len(smpl) / len(lns))
    for i in range(0, amount):
        smplnms.append(smpl[i * len(lns)][sltnlbl])

    print('Neue Probenliste:')
    co = 0
    for i in smplnms:
        print(str(co) + ': ' + i)
        co += 1
# determine detection limit
print('Wenn vorhanden, Blindwertproben zur Bestimmung von NWG/BG auswählen:')
sel = input().split(',')
if sel != ['']:
    for i in lns:
        if 'Ar' in i:
            continue

        blk2 = []
        for j in sel:
            for k in smpl:
                if k[lmnt] == i and k[sltnlbl] == smplnms[int(j)]:
                    blk = []
                    reps = int(k[nmrps])
                    for x in range(0, reps):
                        loc = (x + 1) * -2
                        blk.append(float(k[loc]))
                    blk2.append(np.std(blk))
                    break
        blksd[i] = (sum(blk2) / len(blk2))
else:
    for i in lns:
        if 'Ar' in i:
            continue

        blksd[i] = 9999999999

# get calibration steps
for i in lns:
    stps = []
    if 'Ar' in i:
        continue

    stps.append(i)

    for j in clbrtn:
        if j[lmnt] != i or j[flgs] == '"e"':
            continue
        else:
            stps.append(float(j[cnctrn]))

    cal.append(stps)

# linear regression
for i in cal:
    xs = []
    ys = []
    for j in i[1:]:
        for y in clbrtn:
            reps = int(y[nmrps])
            ntreps = []
            if y[flgs] == '"e"':
                continue
            elif y[lmnt] == i[0] and float(y[cnctrn]) == j:
                for x in range(0, reps):
                    loc = (x + 1) * -2
                    ntreps.append(float(y[loc]))
                ntreps = grubbs(ntreps).tolist()
                for z in range(0, len(ntreps)):
                    xs.append(j)
                ys.extend(ntreps)
                break
    slope, intercept, r, p, std_err = stats.linregress(xs, ys)
    i.append(slope)
    i.append(intercept)
    i.append(r * r)
    i.append((blksd[i[0]] * 3) / slope)
    i.append((blksd[i[0]] * 10) / slope)

# write overview page
overview.write('A1', 'Datum:')
overview.write('A2', 'Startzeit:')
overview.write('A3', 'Endzeit:')
overview.write('A4', 'Anzahl Proben:')
overview.write('A5', 'Anzahl Wellenlängen:')
overview.write('B1', clbrtn[0][dt])
overview.write('B2', clbrtn[0][tm])
overview.write('B3', smpl[-1][tm])
overview.write('B4', amount)
overview.write('B5', len(lns))

overview.write('A7', 'Probenbezeichnung', bold)
overview.write('B7', 'Aufschlussvolumen [ml]', bold)
overview.write('C7', 'Einwaage [g]', bold)
overview.write('D7', 'Verdünnung 1:X, X=', bold)
overview.write('E7', 'Faktor', bold)

overview.set_column('A:D', 20)

row = 7
col = 0

for i in smplnms:
    overview.write(row, col, i)
    overview.write(row, col + 3, 1)
    overview.write(row, col + 4, '=(B' + str(row + 1) + '*D' + str(row + 1) + ')/(1000*C' + str(row + 1) + ')')
    row += 1

# write calibration page
row = 0
col = 0

for i in cal:
    calibrate.write(row, col, 'Elementarlinie', bold)
    calibrate.write(row, col + 1, 'Kalibrierte Konzentrationen [mg/l]', bold)
    calibrate.write(row, col + 2, 'Kalibriergerade', bold)
    calibrate.write(row, col + 3, 'Korrelationskoeffizient R²', bold)
    calibrate.write(row, col + 4, 'NWG [mg/l]', bold)
    calibrate.write(row, col + 5, 'BG [mg/l]', bold)
    calibrate.write(row + 1, col, i[0])
    calibrate.write(row + 1, col + 1, ', '.join(str(e) for e in i[1:-5]))
    calibrate.write(row + 1, col + 2, 'y = ' + '{:.2f}'.format(i[-5]) + 'x + ' + '{:.2f}'.format(i[-4]))
    calibrate.write(row + 1, col + 3, '{:.6f}'.format(i[-3]))
    calibrate.write(row + 1, col + 4, i[-2])
    calibrate.write(row + 1, col + 5, i[-1])
    row += 3

calibrate.set_column('A:A', 20)
calibrate.set_column('B:B', 40)
calibrate.set_column('C:F', 30)

# write results page
row = 2
col = 0

# write list of lines
for i in lns:
    if 'Ar' in i:
        ar.append(lns.index(i))
        continue

    result.write(row, col, i, bold)
    row += 1

row += 1

for i in ar:
    result.write(row, col, lns[i], bold)

# write list of samples
row = 0
col = 1

for i in range(0, amount):
    result.merge_range(row, col, row, col + 3, smplnms[i], bold)
    result.write(row + 1, col, 'Konz. [mg/l]', bold)
    result.write(row + 1, col + 1, 'VI [mg/l]', bold)
    result.write(row + 1, col + 2, 'Gehalt [mg/g]', bold)
    result.write(row + 1, col + 3, 'VI [mg/g]', bold)
    col += 4

result.set_column(0, col + 3, 15)

# write results
row = 2
col = 1

for i in smpl:
    if 'Ar' in i[lmnt]:
        continue
    elif i[flgs] == '"uv"' or i[flgs] == '"x"':
        result.write(row, col, float(i[cnctrn]), red)
        result.write(row, col + 1, i[flgs])
    elif i[flgs] == '"unca"':
        result.write(row, col, i[cnctrn], red)
        result.write(row, col + 1, i[flgs])
    else:
        # calculate concentration
        reps = int(i[nmrps])
        ntreps = []
        for x in range(0, reps):
            loc = (x + 1) * -2
            ntreps.append(float(i[loc]))
        ntreps = grubbs(ntreps).tolist()
        for j in cal:
            if j[0] == i[lmnt]:
                slp = j[-5]
                cpt = j[-4]
                break
        conc = ((sum(ntreps) / len(ntreps)) - cpt) / slp
        # calculate confidence interval
        vi = ((np.std(ntreps) / slp) * stats.t.ppf(1 - 0.025, len(ntreps) - 1)) / math.sqrt(len(ntreps))
        result.write(row, col, conc)
        result.write(row, col + 1, vi)
        result.write(row, col + 2, '=' + xl_rowcol_to_cell(row, col) + '*Overview!E' + str(int(((col + 1) / 4) + 8)))
        result.write(row, col + 3,
                     '=' + xl_rowcol_to_cell(row, col + 1) + '*Overview!E' + str(int(((col + 1) / 4) + 8)))

    if row == len(lns) - 1 + len(ar):
        result.write(row + 1, col, 'Verhältnis', bold)
        row = 2
        col += 4
    else:
        row += 1

# write argon lines
row = 3 + len(lns) - len(ar)
col = 1

for i in smpl:
    if not 'Ar' in i[lmnt]:
        continue
    else:
        result.write(row, col, float(i[nt]) / float(arint[i[lmnt]]))

    if row == 2 + len(lns):
        row = 3 + len(lns) - len(ar)
        col += 4
    else:
        row += 1

workbook.close()

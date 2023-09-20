#!/usr/bin/env python3
#
# Doel:
# Dit script leest een CSV-file van de Rabobank in en maakt hem klaar voor Excel.
# Het script voert de volgende bewerkingen uit:
# - Splitst de CSV op in een apart bestand per rekening.
# - Rubriceert alle mutaties aan de hand van een tabel met regular expressions.
#   Records die niet matchen met een regex worden gelogd in de logfile.
# - Sorteert de mutaties op rubriek, relatienaam en volgnummer.
# - Voegt formules toe die de totaaltellingen van de rubrieken doen.
# - Schrijft een nieuw CSV-bestand voor elke rekening in een formaat dat helpt bij de belastingaangifte.
#
# Starten als volgt:
# ./rabo-csv.py --infile infile.csv \
#               --matchfile matchfile.csv \
#               --verbosity 2 \
#               --logfile logfile.log
#
# Optioneel kun je ook een output-directory meegeven via de optie --outdir
#
# (c) 2023 TvdP

import os, sys, re, csv, chardet, argparse, datetime, locale, operator

def logmsg(severity, message):
    if log_open:
        if   severity <= args.verbosity:
            if   severity == LOG_INFO:
                svrty = ' INFO '
            elif severity == LOG_WARN:
                svrty = ' WARN '
            elif severity == LOG_ERROR:
                svrty = ' ERROR '
            elif severity == LOG_DEBUG:
                svrty = ' DEBUG '
            else:
                svrty = ' UNKOWN '
            try:
                logfile.write(datetime.datetime.today().strftime("%Y%m%d%H%M%S") + svrty + message + '\n')
            except (OSError, IOError) as e:
                print(e, file=sys.stderr)
                sys.exit(2)


def get_accounts_from_csv_file(csv_infile_dict_list):
    list_of_accounts = []
    for row in csv_infile_dict_list:
        if row['IBAN/BBAN'] not in list_of_accounts:
            list_of_accounts.append(row['IBAN/BBAN'])
    return list_of_accounts


def match_category(row, matchfile_dict_list):
    line = 0
    for matching_rule in matchfile_dict_list:
        line += 1
        try:
            if re.search(matching_rule['Regex'], row[matching_rule['Column']]):
                logmsg(LOG_INFO, "Matched '" + matching_rule['Category'] +
                                    "' on '" + matching_rule['Regex'] +
                                    "' in '" + matching_rule['Column'] +
                                    "' using matching rule in line " + str(line+1))
                return matching_rule['Category']
        except re.error as e:
            logmsg(LOG_ERROR, "Invalid regular expression in line " + str(line+1) + " of the matchfile: " + str(e))
            sys.exit(2)

    logmsg(LOG_WARN, "No match on volgnummer " + '{0:>6}'.format(str(int(row['Volgnr']))) +
                                   ", tegenrekening: " +                 row['Tegenrekening IBAN/BBAN'] +
                                   ", bedrag: "        + '{0:>8}'.format(row['Bedrag']) +
                                   ", tegenpartij '"   +                 row['Naam tegenpartij'] +
                                  "', regarding '"     +                 row['Omschrijving-1'] + "'")
    return ''


def add_spreadsheet_formulas(row, first_row, last_row):
    value_bij = '=SUM(F{0}:F{1})'.format(first_row, last_row)
    value_af  = '=SUM(G{0}:G{1})'.format(first_row, last_row)
    row.update({'Rubriek bij': value_bij})
    row.update({'Rubriek af':  value_af})


def write_csv_file(csv_infile_dict_list, matchfile_dict_list, account_nr, outdir):
    try:
        csv_output_file = outdir + '/' + account_nr + '.csv'
        logmsg(LOG_INFO, "Opening csv_output_file " + csv_output_file)
        csv_outfile = open(csv_output_file, 'w', newline='')
        fieldnames = ['Nr', 'Boekdatum', 'Rentedatum', 'Rekening (IBAN)', 'Rubriek', 'Bedrag bij', 'Bedrag af', 'Rubriek bij', 'Rubriek af', 'Saldo', 'Omschrijving', 'Tegenrekening (IBAN)', 'Relatienaam']
        csv_outfile_writer = csv.DictWriter(csv_outfile, fieldnames=fieldnames)
        csv_outfile_writer.writeheader()
        csv_outfile_dict_list = []

        # Build a list of dicts for the output CSV in memory.
        # This list will NOT contain headers; those were already written to the file a few lines back.
        rowcnt = 0
        for row in csv_infile_dict_list:
            if row['IBAN/BBAN'] == account_nr:
                rowcnt += 1
                bedrag_bij = locale.currency(0, symbol=False) if locale.atof(row['Bedrag']) <  0 else locale.currency( locale.atof(row['Bedrag']), symbol=False)
                bedrag_af  = locale.currency(0, symbol=False) if locale.atof(row['Bedrag']) >= 0 else locale.currency(-locale.atof(row['Bedrag']), symbol=False)
                csv_outfile_dict_list.append({
                    'Nr':                   rowcnt,
                    'Boekdatum':            row['Datum'],
                    'Rentedatum':           row['Rentedatum'],
                    'Rekening (IBAN)':      row['IBAN/BBAN'],
                    'Rubriek':              match_category(row, matchfile_dict_list),
                    'Bedrag bij':           bedrag_bij,
                    'Bedrag af':            bedrag_af,
                    'Rubriek bij':          '',
                    'Rubriek af':           '',
                    'Saldo':                row['Saldo na trn'],
                    'Omschrijving':         row['Omschrijving-1'],
                    'Tegenrekening (IBAN)': row['Tegenrekening IBAN/BBAN'],
                    'Relatienaam':          row['Naam tegenpartij']
                })

        csv_outfile_dict_list_sorted = sorted(csv_outfile_dict_list, key=operator.itemgetter('Rubriek',
                                                                                             'Relatienaam',
                                                                                             'Nr'))
        this_row = 0
        category_start = 0
        this_category = ''
        next_category = ''
        for row in csv_outfile_dict_list_sorted:
            this_category = row['Rubriek']
            if this_row < rowcnt-1:
                next_category = csv_outfile_dict_list_sorted[this_row+1]['Rubriek']
            else:
                next_category = ''
            if this_category == next_category:
                logmsg(LOG_DEBUG, "Writing row " + str(this_row) + ", Nr: " + str(row['Nr']) + ", rubriek: " + row['Rubriek'])
                csv_outfile_writer.writerow(row)
            else:
                add_spreadsheet_formulas(row, category_start+2, this_row+2)
                logmsg(LOG_DEBUG, "Writing row " + str(this_row) + ", Nr: " + str(row['Nr']) + ", rubriek: " + row['Rubriek'] + ", Rubriek bij: " + row['Rubriek bij'] + ", Rubriek af: " + row['Rubriek af'])
                csv_outfile_writer.writerow(row)
                category_start = this_row+1
            this_row += 1

    except csv.Error as e:
        logmsg(LOG_ERROR, "outfile: " + csv_output_file + ", line: " + csv_outfile_writer.line_num + ", errmsg: " + e)
        sys.exit(2)

    except (OSError, IOError) as e:
        logmsg(LOG_ERROR, e)
        sys.exit(2)

# Header layout from the Rabobank CSV:
# ['IBAN/BBAN', 'Munt', 'BIC', 'Volgnr', 'Datum', 'Rentedatum', 'Bedrag', 'Saldo na trn', 'Tegenrekening IBAN/BBAN', 'Naam tegenpartij', 'Naam uiteindelijke partij', 'Naam initiërende partij', 'BIC tegenpartij', 'Code', 'Batch ID', 'Transactiereferentie', 'Machtigingskenmerk', 'Incassant ID', 'Betalingskenmerk', 'Omschrijving-1', 'Omschrijving-2', 'Omschrijving-3', 'Reden retour', 'Oorspr bedrag', 'Oorspr munt', 'Koers']

locale.setlocale(locale.LC_ALL, locale.getlocale())

LOG_DEBUG = 4
LOG_INFO  = 3
LOG_WARN  = 2
LOG_ERROR = 1

parser = argparse.ArgumentParser(description='This program reads a Rabobank CSV, splits it into separate accounts, categorizes the transactions and converts the output to a format suitable to support the annual income tax declaration in a spreadsheet.')
parser.add_argument('--infile',          required=True)  # The CSV-file from the Rabobank
parser.add_argument('--matchfile',       required=True)  # a CSV-file containing the matching expressions and related categories
parser.add_argument('--outdir',          default=None)   # The output directory
parser.add_argument('--logfile',         default=sys.stderr)
parser.add_argument('--verbosity',       type=int, choices=[1,2,3,4], default=LOG_WARN)
args = parser.parse_args()

if args.outdir == None:
    outdir = os.path.dirname(os.path.abspath(args.infile))
else:
    outdir = args.outdir

csv_line = 0

# Open log file
log_open = False
if args.logfile == sys.stderr:
    logfile = args.logfile
    log_open = True
else:
    try:
        logfile = open(args.logfile, 'a')
        log_open = True
    except (OSError, IOError) as e:
        print(e, file=sys.stderr)
        sys.exit(2)

# Some sanity checks
if not os.path.isfile(args.infile):
    logmsg(LOG_ERROR, args.infile + ": input file does not exist")
    sys.exit(2)

if not os.path.isfile(args.matchfile):
    logmsg(LOG_ERROR, args.matchfile + ": input file does not exist")
    sys.exit(2)

if not os.path.isdir(outdir):
    logmsg(LOG_ERROR, outdir + ": output directory does not exist")
    sys.exit(2)

# Determine the encoding of the infile and the matchfile
try:
    with open(args.infile, 'rb') as rawdata:
        chardet_result = chardet.detect(rawdata.read(100000))
        logmsg(LOG_INFO, "Encoding of file " + args.infile + "is '" + chardet_result['encoding'] + "' with confidence " + str(chardet_result['confidence']))
        infile_encoding = chardet_result['encoding']
        rawdata.close()
except (OSError, IOError) as e:
    logmsg(LOG_ERROR, e)
    sys.exit(2)

try:
    with open(args.matchfile, 'rb') as rawdata:
        chardet_result = chardet.detect(rawdata.read(100000))
        logmsg(LOG_INFO, "Encoding of file " + args.matchfile + "is '" + chardet_result['encoding'] + "' with confidence " + str(chardet_result['confidence']))
        matchfile_encoding = chardet_result['encoding']
        rawdata.close()
except (OSError, IOError) as e:
    logmsg(LOG_ERROR, e)
    sys.exit(2)

# Open the input files and read them into memory
try:
    with open(args.matchfile, 'r', encoding=matchfile_encoding) as matchfile:
        matchreader = csv.DictReader(matchfile, delimiter=',', quotechar='"')
        matchfile_dict_list = list(matchreader)
        matchfile.close()
except (OSError, IOError) as e:
    logmsg(LOG_ERROR, e)
    sys.exit(2)

try:
    with open(args.infile, 'r', encoding=infile_encoding) as csv_infile:
        csv_infile_reader = csv.DictReader(csv_infile, delimiter=',', quotechar='"')
        csv_infile_dict_list = list(csv_infile_reader)
        csv_infile.close()
except (OSError, IOError) as e:
    logmsg(LOG_ERROR, e)
    sys.exit(2)

# Do the work
account_list = get_accounts_from_csv_file(csv_infile_dict_list)
for account_nr in account_list:
    write_csv_file(csv_infile_dict_list, matchfile_dict_list, account_nr, outdir)

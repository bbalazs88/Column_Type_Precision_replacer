import pandas as pd
import os

# file_loc = 'c:\Tesztfajlok\RDR_CLIENT_PROFITABILITY.xls'
# itthon: c:\simple.dev\Mapek\CDR_CORRECTION_TEMPLATE_FINAL.xlsx - mukodik!
# p és r a két oszlop CDR-nel.
# c:\simple.dev\Mapek\DW_CUSTOMER_MONITORING_ATTR2.xlsx
# n - o - p oszlopok, kulon van a precision

result = ''
sql_type = []
precision = []

while True:
    try:
        file_loc = 'c:\simple.dev\Mapek\DW_CUSTOMER_MONITORING_ATTR2.xlsx' #input("File teljes elérési útvonala: ")
        separate_types_prec = input("Külön van a típus a hossztól? Y/N: ").upper()

        while separate_types_prec not in ("Y", "N"):
            print("Csak Y vagy N fogadható el.\n")
            separate_types_prec = input("Külön van a típus a hossztól? Y/N: ").upper()

        if separate_types_prec == 'N':
            col_letter = input("Add meg a betűjelét az oszlopneveket tartalmazó oszlopnak: ")
            type_letter = input("Add meg a betűjelét a típusokat tartalmazó oszlopnak: ")

            cols = pd.read_excel(file_loc, usecols=col_letter.upper())
            types = pd.read_excel(file_loc, usecols=type_letter.upper(), dtype=str).fillna(0)

            cols_list_old = cols.values.tolist()
            cols_list = [item for sublist in cols_list_old for item in sublist]

            types_list_old = types.values.tolist()
            types_list = [item for sublist in types_list_old for item in sublist]

        elif separate_types_prec == 'Y':
            col_letter = input("Add meg a betűjelét az oszlopneveket tartalmazó oszlopnak: ").upper()
            type_letter = input("Add meg a betűjelét a típusokat tartalmazó oszlopnak: ").upper()
            prec_letter = input("Add meg a betűjelét a hosszt tartalmazó oszlopnak: ").upper()

            cols = pd.read_excel(file_loc, usecols=col_letter)
            types = pd.read_excel(file_loc, usecols=type_letter, dtype=str).fillna(0)
            prec = pd.read_excel(file_loc, usecols=prec_letter, skip_blank_lines=False, dtype=str)

            cols_list_old = cols.values.tolist()
            cols_list = [item for sublist in cols_list_old for item in sublist]

            types_list_old = types.values.tolist()
            types_list = [item for sublist in types_list_old for item in sublist]

            prec_list_old = prec.index.tolist()
            prec_string = []
            prec_list = []

            for i in prec_list_old:
                prec_string.append(str(i))

            for i in prec_string:
                if i == 'nan':
                    prec_list.append(0)
                else:
                    prec_list.append(float(i))

            prec_list = list(map(int, prec_list))

    except:
        continue

    break

if separate_types_prec == 'Y':
    print('************************************************************** \n')
    print("Oszlopok: \n", cols_list, "\n")
    print("Adattípusok: \n", types_list, "\n")
    print("Hosszok: \n", prec_list, "\n")

    drop_cols = int(input("Hányadik elemnél kezdődik az oszlopok felsorolása: "))
    cols_list = cols_list[drop_cols - 1:]
    cols_list = [x.upper() for x in cols_list]

    types_list = types_list[drop_cols - 1:]
    types_list = [x.upper() for x in types_list]

    prec_list = prec_list[drop_cols - 1:]
    prec_list = [x for x in prec_list]

    print('************************************************************** \n')
    print("A végső sorrend:")
    print(cols_list, types_list, prec_list, "\n", sep="\n")

    for i in types_list:

        if i == "BIGINT":
            sql_type.append(-5)

        if i == "DATE":
            sql_type.append(9)

        if i == "VARCHAR":
            sql_type.append(12)

        if i == "DOUBLE":
            sql_type.append(8)

        if i == "INTEGER":
            sql_type.append(4)

    for i, j, k in zip(cols_list, sql_type, prec_list):
        result += """BEGIN DSSUBRECORD
             Name "%s"
             SqlType "%d"
             Precision "%d"
             Scale "0"
             Nullable "0"
             KeyPosition "0"
             DisplaySize "0"
             Group "0"
             SortKey "0"
             SortType "0"
             AllowCRLF "0"
             LevelNo "0"
             Occurs "0"
             PadNulls "0"
             SignOption "0"
             SortingOrder "0"
             ArrayHandling "0"
             SyncIndicator "0"
             PadChar ""
             ExtendedPrecision "0"
             TaggedSubrec "0"
             OccursVarying "0"
             PKeyIsCaseless "0"
             SCDPurpose "0"
          END DSSUBRECORD \n""" % (i, j, k)

    with open(os.environ['USERPROFILE'] + "\Documents\dsx.txt", "w") as text_file:
        text_file.write(result)

    input("Kész. Keresd a dsx.txt filet a Documents mappádban és nyomj egy gombot a kilépéshez.")

else:
    print('************************************************************** \n')
    print("Oszlopok: \n", cols_list, "\n")
    print("Adattípusok: \n", types_list, "\n")

    drop_cols = int(input("Hányadik elemnél kezdődik az oszlopok felsorolása: "))
    cols_list = cols_list[drop_cols - 1:]
    cols_list = [x.upper() for x in cols_list]

    types_list = types_list[drop_cols - 1:]
    types_list = [x.upper() for x in types_list]

    print('************************************************************** \n')
    print("A végső sorrend: ")
    print(cols_list, types_list, sep="\n")

    for i in types_list:

        if i == "BIGINT":
            sql_type.append(-5)
            precision.append(0)

        if i == "DATE":
            sql_type.append(9)
            precision.append(0)

        if i == "VARCHAR(30)":
            sql_type.append(12)
            precision.append(30)

        if i == "VARCHAR(255)":
            sql_type.append(12)
            precision.append(255)

        if i == "VARCHAR(1)":
            sql_type.append(12)
            precision.append(1)

        if i == "DOUBLE":
            sql_type.append(8)
            precision.append(0)

        if i == "INTEGER":
            sql_type.append(4)
            precision.append(0)

    for i, j, k in zip(cols_list, sql_type, precision):
        result += """BEGIN DSSUBRECORD
             Name "%s"
             SqlType "%d"
             Precision "%d"
             Scale "0"
             Nullable "0"
             KeyPosition "0"
             DisplaySize "0"
             Group "0"
             SortKey "0"
             SortType "0"
             AllowCRLF "0"
             LevelNo "0"
             Occurs "0"
             PadNulls "0"
             SignOption "0"
             SortingOrder "0"
             ArrayHandling "0"
             SyncIndicator "0"
             PadChar ""
             ExtendedPrecision "0"
             TaggedSubrec "0"
             OccursVarying "0"
             PKeyIsCaseless "0"
             SCDPurpose "0"
          END DSSUBRECORD \n""" % (i, j, k)

    with open(os.environ['USERPROFILE'] + "\Documents\dsx.txt", "w") as text_file:
        text_file.write(result)

    input("Kész. Keresd a dsx.txt filet a Documents mappádban és nyomj egy gombot a kilépéshez.")
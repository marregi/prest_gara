import xlrd

hil_izenak = {
    1: 'Urtarrila',
    2: 'Otsaila',
    3: 'Martxoa',
    4: 'Apirila',
    5: 'Maiatza',
    6: 'Ekaina',
    7: 'Uztaila',
    8: 'Abuztua',
    9: 'Iraila',
    10: 'Urria',
    11: 'Azaroa',
    12: 'Abendua',
}

def hileko_gordina_garbia(salmenta_lerroak):
    urte_hil_irabazi_gordina = {}
    urte_hil_irabazi_garbia = {}

    for lerroa in salmenta_lerroak:
        urtea = lerroa['urtea']
        hila = lerroa['hila']
        if urtea in urte_hil_irabazi_garbia:
            if hila in urte_hil_irabazi_garbia[urtea]:
                urte_hil_irabazi_gordina[urtea][hila] += lerroa['totala']
                urte_hil_irabazi_garbia[urtea][hila] += lerroa['irabazia']
            else:
                urte_hil_irabazi_gordina[urtea][hila] = lerroa['totala']
                urte_hil_irabazi_garbia[urtea][hila] = lerroa['irabazia']
        else:
            urte_hil_irabazi_gordina[urtea] = {hila: lerroa['totala']}
            urte_hil_irabazi_garbia[urtea] = {hila: lerroa['irabazia']}
    return urte_hil_irabazi_gordina, urte_hil_irabazi_garbia


def urteko_errentagarriena(salmenta_lerroak):
    urte_produktu_irabazi_gordina = {}
    urte_produktu_irabazi_garbia = {}

    for lerroa in salmenta_lerroak:
        urtea = f"{lerroa['urtea']}"
        kodea = f"{lerroa['produktu_kodea']}"
        if urtea in urte_produktu_irabazi_garbia:
            if kodea in urte_produktu_irabazi_gordina[urtea]:
                urte_produktu_irabazi_gordina[urtea][kodea] += lerroa['totala']
                urte_produktu_irabazi_garbia[urtea][kodea] += lerroa[
                    'irabazia']
            else:
                urte_produktu_irabazi_gordina[urtea][kodea] = lerroa['totala']
                urte_produktu_irabazi_garbia[urtea][kodea] = lerroa['irabazia']
        else:
            urte_produktu_irabazi_gordina[urtea] = {kodea: lerroa['totala']}
            urte_produktu_irabazi_garbia[urtea] = {kodea: lerroa['irabazia']}
    return urte_produktu_irabazi_gordina, urte_produktu_irabazi_garbia


# Urte-Hileko Produkturik Salduena eta Errentagarriena


def hileko_errentagarriena(salmenta_lerroak):
    urte_hil_produktu_irabazi_gordina = {}
    urte_hil_produktu_irabazi_garbia = {}

    for lerroa in salmenta_lerroak:
        urtea = lerroa['urtea']
        hila = lerroa['hila']
        kodea = lerroa['produktu_kodea']
        if urtea in urte_hil_produktu_irabazi_garbia:
            if hila in urte_hil_produktu_irabazi_gordina[urtea]:
                if kodea in urte_hil_produktu_irabazi_gordina[urtea][hila]:
                    urte_hil_produktu_irabazi_gordina[urtea][hila][kodea] += \
                        lerroa['totala']
                    urte_hil_produktu_irabazi_garbia[urtea][hila][kodea] += \
                        lerroa['irabazia']
                else:
                    urte_hil_produktu_irabazi_gordina[urtea][hila][kodea] = \
                        lerroa['totala']
                    urte_hil_produktu_irabazi_garbia[urtea][hila][kodea] = \
                        lerroa['irabazia']
            else:
                urte_hil_produktu_irabazi_gordina[urtea][hila] = {kodea: lerroa['totala']}
                urte_hil_produktu_irabazi_garbia[urtea][hila] = {kodea: lerroa['irabazia']}
        else:
            urte_hil_produktu_irabazi_gordina[urtea] = {hila: {kodea: lerroa[
                'totala']}}
            urte_hil_produktu_irabazi_garbia[urtea] = {hila: {kodea: lerroa[
                'irabazia']}}
    return urte_hil_produktu_irabazi_gordina, urte_hil_produktu_irabazi_garbia


def produktu_salmenta(kodea, salmenta_lerroak):
    produktua_hileko_kantitatea = {}
    for lerroa in salmenta_lerroak:
        if lerroa['produktu_kodea'] == kodea:
            urtea = lerroa['urtea']
            hila = lerroa['hila']
            if urtea in produktua_hileko_kantitatea:
                if hila in produktua_hileko_kantitatea[urtea]:
                    produktua_hileko_kantitatea[urtea][hila] += lerroa[
                        'kantitatea']
                else:
                    produktua_hileko_kantitatea[urtea][hila] = lerroa[
                        'kantitatea']
            else:
                produktua_hileko_kantitatea[urtea] = {hila: lerroa[
                    'kantitatea']}
    return produktua_hileko_kantitatea


def irakurri_produktuak():
    wb = xlrd.open_workbook('denda.xlsx')
    produktuak = wb.sheet_by_name('Produktuak')

    produktukoak = {}

    for lerroa in range(produktuak.nrows - 1):
        produktu_id = lerroa + 1
        produktukoak[produktu_id] = produktuak.cell_value(lerroa + 1, 1)
    return produktukoak


def irakurri_salmentak():
    wb = xlrd.open_workbook('denda.xlsx')
    salmentak = wb.sheet_by_name('Salmenta Orria')
    salmenta_lerroak = []
    for lerroa in range(salmentak.nrows - 1):
        salmenta_lerro = {}
        salmenta_lerro["urtea"] = int(salmentak.cell_value(lerroa + 1, 0))
        salmenta_lerro["hila"] = int(salmentak.cell_value(lerroa + 1, 1))
        salmenta_lerro["eguna"] = int(salmentak.cell_value(lerroa + 1, 2))
        salmenta_lerro["produktu_kodea"] = int(salmentak.cell_value(lerroa +
                                                                  1, 3))
        salmenta_lerro["kantitatea"] = salmentak.cell_value(lerroa + 1, 4)
        salmenta_lerro["zenbatekoa"] = salmentak.cell_value(lerroa + 1, 5)
        salmenta_lerro["totala"] = salmentak.cell_value(lerroa + 1, 6)
        salmenta_lerro["irabazia"] = salmentak.cell_value(lerroa + 1, 7)
        salmenta_lerroak.append(salmenta_lerro)
    return salmenta_lerroak


def salmenta_estatistikak():
    salmenta_lerroak = irakurri_salmentak()
    produktuak = irakurri_produktuak()
    testua = """Zein estatistika lortu nahi duzu?
            1. Urteko eta hilabeteko irabazi gordina/garbia
            2. Urte eta hileko produkturik errentagarriena (garbia/gordina)
            3. Produktu baten salmenta kantitatea hileko
            Sartu zenbakia: """
    estatistika_kodea = int(input(testua))
    while estatistika_kodea not in [1, 2, 3]:
        print("Kodea ez da zuzena")
        estatistika_kodea = int(input(testua))
    if estatistika_kodea == 1:
        hileko_gordina, hileko_garbia = hileko_gordina_garbia(salmenta_lerroak)
        for urtea in hileko_gordina.keys():
            urteko_gordina = 0
            urteko_garbia = 0
            for hila in hileko_garbia[urtea].keys():
                urteko_gordina += hileko_gordina[urtea][hila]
                urteko_garbia += hileko_garbia[urtea][hila]
                print(f"{urtea}-ko {hil_izenak[hila]}-an gordina "
                      f"{hileko_gordina[urtea][hila]} izan da eta garbia "
                      f"{hileko_gordina[urtea][hila]}")
            print(f"Urteko gordina {urteko_gordina} izan da eta garbia {urteko_garbia}")
    elif estatistika_kodea == 2:
        hileko_errentagarriena_gordina, hileko_errentagarriena_garbia = hileko_errentagarriena(
            salmenta_lerroak)
        for urtea in hileko_errentagarriena_gordina.keys():
            for hila in hileko_errentagarriena_gordina[urtea].keys():
                gordina = max(hileko_errentagarriena_gordina[urtea][hila],
                              key=hileko_errentagarriena_gordina[urtea][hila].get)
                garbia = max(hileko_errentagarriena_garbia[urtea][hila],
                             key=hileko_errentagarriena_garbia[urtea][hila].get)
                print(
                    f"{urtea}-ko {hil_izenak[hila]}-ean {produktuak[gordina]} izan da "
                    f"irabazi gordin handiena izan duena")
                print(
                    f"{urtea}-ko {hil_izenak[hila]}-ean {produktuak[garbia]} izan da "
                    f"irabazi garbi handiena izan duena")
    else:
        produktu_kodea = int(input("Sartu produktu kodea: "))
        while produktu_kodea not in produktuak.keys():
            produktu_kodea = input(
                "Produktu kodea ez da existitzen sartu beste bat: ")
        produktu_kantitatea_hileko = produktu_salmenta(
            produktu_kodea, salmenta_lerroak)
        # hileko_kantitatea = max(produktu_kantitatea_hileko,
        #                         key=produktu_kantitatea_hileko.get)
        for urtea in produktu_kantitatea_hileko.keys():
            for hila in produktu_kantitatea_hileko[urtea].keys():
                print(f"{urtea}-ko {hil_izenak[hila]}-n "
                      f"{produktu_kantitatea_hileko[urtea][hila]} unitate "
                      f"saldu dira")
            # print(
            #     f"{urteko_kantitatea} izan da produktua gehien saldu den urtea eta {produktua_urteko_kantitatea[urteko_kantitatea]} unitate saldu dira")
            # print(
            #     f"{hileko_kantitatea.split('-')[0]}-ko {hileko_kantitatea.split('-')[1]}-n izan da produktua gehien saldu den hila eta {produktua_hileko_kantitatea[hileko_kantitatea]} unitate saldu dira")


if __name__ == '__main__':
    salmenta_estatistikak()

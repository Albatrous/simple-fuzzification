import csv
from colorama import init, Fore, Style
import xlwt


def fuzzification(efektivitas, kualitas):
    """Convert crisp input to fuzzy input based on the rules"""
    efektivitas = int(efektivitas)
    kualitas = int(kualitas)

    # EFEKTIVITAS
    buruk = 0
    biasa = 0
    baik = 0
    if efektivitas <= 10:
        buruk = 1
    elif efektivitas in range(11, 30):
        buruk = (-(efektivitas - 30)) / (30 - 10)  # -(x - c)/(c - b)
        biasa = (efektivitas - 10) / (30 - 10)  # (x - a)/(b - a)
    elif efektivitas in range(30, 51):
        biasa = 1
    elif efektivitas in range(51, 70):
        biasa = (-(efektivitas - 70)) / (70 - 50)  # -(x - d)/(d - c)
        baik = (efektivitas - 50) / (70 - 50)  # (x - a)/(b - a)
    else:
        baik = 1

    # KUALITAS
    sangat_rendah = 0
    rendah = 0
    bagus = 0
    sangat_bagus = 0
    if kualitas <= 15:
        sangat_rendah = 1
    elif kualitas in range(16, 30):
        sangat_rendah = (-(kualitas - 30)) / (30 - 15)  # -(x - c)/(c - b)
        rendah = (kualitas - 15) / (30 - 15)  # (x - a)/(b - a)
    elif kualitas in range(30, 46):
        rendah = 1
    elif kualitas in range(46, 60):
        rendah = (-(kualitas - 60)) / (60 - 45)  # -(x - d)/(d - c)
        bagus = (kualitas - 45) / (60 - 45)  # (x - a)/(b - a)
    elif kualitas in range(60, 76):
        bagus = 1
    elif kualitas in range(76, 90):
        bagus = (-(kualitas - 90)) / (90 - 75)  # -(x - d)/(d - c)
        sangat_bagus = (kualitas - 75) / (90 - 75)  # (x - a)/(b - a)
    else:
        sangat_bagus = 1

    fuz_input = {
        "efektivitas": {"buruk": buruk, "biasa": biasa, "baik": baik},
        "kualitas": {
            "sangat_rendah": sangat_rendah,
            "rendah": rendah,
            "bagus": bagus,
            "sangat_bagus": sangat_bagus,
        },
    }
    return fuz_input


def inferention(fuz_input):
    """Convert fuzzy input to fuzzy output"""
    # Nilai Kelayakan
    worth_value = {}

    def inferention_low(efektivitas, kualitas):
        if (efektivitas != 0) and (kualitas != 0):
            fuz_output = {"Low": min(efektivitas, kualitas)}
            try:
                if worth_value["Low"] < fuz_output["Low"]:
                    worth_value.update(fuz_output)
            except KeyError:
                worth_value.update(fuz_output)

    def inferention_normal(efektivitas, kualitas):
        if (efektivitas != 0) and (kualitas != 0):
            fuz_output = {"Normal": min(efektivitas, kualitas)}
            try:
                if worth_value["Normal"] < fuz_output["Normal"]:
                    worth_value.update(fuz_output)
            except KeyError:
                worth_value.update(fuz_output)

    def inferention_high(efektivitas, kualitas):
        if (efektivitas != 0) and (kualitas != 0):
            fuz_output = {"High": min(efektivitas, kualitas)}
            try:
                if worth_value["High"] < fuz_output["High"]:
                    worth_value.update(fuz_output)
            except KeyError:
                worth_value.update(fuz_output)

    # Kategori Efektivitas
    kat_efektif = fuz_input["efektivitas"]
    # Kategori Kualitas
    kat_kualiti = fuz_input["kualitas"]

    # The RULES
    inferention_low(kat_efektif["buruk"], kat_kualiti["sangat_rendah"])
    inferention_low(kat_efektif["biasa"], kat_kualiti["sangat_rendah"])
    inferention_low(kat_efektif["buruk"], kat_kualiti["rendah"])
    inferention_low(kat_efektif["buruk"], kat_kualiti["bagus"])
    inferention_normal(kat_efektif["baik"], kat_kualiti["sangat_rendah"])
    inferention_normal(kat_efektif["biasa"], kat_kualiti["rendah"])
    inferention_normal(kat_efektif["baik"], kat_kualiti["rendah"])
    inferention_normal(kat_efektif["buruk"], kat_kualiti["sangat_bagus"])
    inferention_high(kat_efektif["biasa"], kat_kualiti["bagus"])
    inferention_high(kat_efektif["baik"], kat_kualiti["bagus"])
    inferention_high(kat_efektif["biasa"], kat_kualiti["sangat_bagus"])
    inferention_high(kat_efektif["baik"], kat_kualiti["sangat_bagus"])

    return worth_value


def defuzzification(worth_value):
    """Defuzzification to make crisp output"""
    try:
        low = worth_value["Low"]
    except KeyError:
        low = 0

    try:
        Normal = worth_value["Normal"]
    except KeyError:
        Normal = 0

    try:
        High = worth_value["High"]
    except KeyError:
        High = 0

    crisp_output = ((low * 50) + (Normal * 75) + (High * 100)) / (low + Normal + High)

    return crisp_output


def main():
    init(autoreset=True)
    csv_file = open("karyawan.csv", "r")
    dataset = csv.DictReader(csv_file)
    results = []

    for data in dataset:
        fuzzi = fuzzification(data["efektivitas"], data["kualitas"])
        fuz_rules = inferention(fuzzi)
        crisp_output = defuzzification(fuz_rules)
        results.append(
            {
                "id": data["id"],
                "data": {
                    "fuzzification": fuzzi,
                    "worth_value": fuz_rules,
                    "crisp_output": crisp_output,
                },
            }
        )

    results.sort(key=lambda x: x["result"]["crisp_output"], reverse=True)

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Rank")

    j = 0
    ws.write(0, 0, "Employee ID")

    for data in results[:10]:
        # Save to excel
        j += 1
        ws.write(j, 0, int(data["id"]))

        # Print result efektivitas
        efektivitas = data["data"]["fuzzification"]["efektivitas"]
        efektivitas = [
            f"[{Fore.RED}>{Style.RESET_ALL}] {i.capitalize()} = {efektivitas[i]}"
            for i in efektivitas
            if efektivitas[i] != 0
        ]

        efektivitas = "\n".join(efektivitas)

        # Print result kualitas
        kualitas = data["data"]["fuzzification"]["kualitas"]
        kualitas = [
            f"[{Fore.RED}>{Style.RESET_ALL}] {i.capitalize()} = {kualitas[i]}"
            for i in kualitas
            if kualitas[i] != 0
        ]
        kualitas = "\n".join(kualitas)

        # Print result of worth value
        worth_value = data["data"]["worth_value"]
        worth_value = [
            f"[{Fore.RED}>{Style.RESET_ALL}] {i} = {worth_value[i]}"
            for i in worth_value
        ]
        worth_value = "\n".join(worth_value)
        crisp_output = data["data"]["crisp_output"]
        header = f"========== {Fore.LIGHTCYAN_EX}ID: {data['id']}{Style.RESET_ALL} =========="
        footer = "=" * (len(header) - 9)

        # The OUTPUT
        kata = (
            f"{header}{Style.RESET_ALL}\n"
            f"[{Fore.LIGHTYELLOW_EX}*{Style.RESET_ALL}] {Fore.LIGHTYELLOW_EX}FUZZIFICATION{Style.RESET_ALL}\n"
            f"[{Fore.BLUE}#{Style.RESET_ALL}] {Fore.BLUE}Efektivitas:{Style.RESET_ALL}\n"
            f"{efektivitas}\n"
            f"[{Fore.BLUE}#{Style.RESET_ALL}] {Fore.BLUE}Kualitas:{Style.RESET_ALL}\n"
            f"{kualitas}\n\n"
            f"[{Fore.LIGHTYELLOW_EX}*{Style.RESET_ALL}] {Fore.LIGHTYELLOW_EX}WORTH VALUE{Style.RESET_ALL}\n"
            f"{worth_value}\n\n"
            f"[{Fore.LIGHTYELLOW_EX}*{Style.RESET_ALL}] {Fore.LIGHTYELLOW_EX}CRISP OUTPUT{Style.RESET_ALL} = {crisp_output}\n"
            f"{footer}"
        )

        print(kata, end="\n\n")

    # Save all of the top 10 employee id to xls file
    wb.save("Top 10 Karyawan.xls")


if __name__ == "__main__":
    main()

def indian_number_to_words(number):
    try:
        n = int(round(float(str(number).replace(',', '').replace('₹', '').replace('Rs.', '').replace('/-', '').strip())))
    except Exception:
        return str(number)

    if n == 0:
        return "Zero"

    units = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
             "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
             "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def two_digits(num):
        if num < 20:
            return units[num]
        else:
            return tens[num // 10] + (" " + units[num % 10] if num % 10 != 0 else "")

    def three_digits(num):
        hundred = num // 100
        rem = num % 100
        res = ""
        if hundred > 0:
            res += units[hundred] + " Hundred"
            if rem > 0:
                res += " and "
        if rem > 0:
            res += two_digits(rem)
        return res

    parts = []
    crore = n // 10000000
    rem = n % 10000000
    if crore > 0:
        parts.append(f"{three_digits(crore)} Crore")

    lakh = rem // 100000
    rem = rem % 100000
    if lakh > 0:
        parts.append(f"{two_digits(lakh)} Lakh")

    thousand = rem // 1000
    rem = rem % 1000
    if thousand > 0:
        parts.append(f"{two_digits(thousand)} Thousand")

    if rem > 0:
        parts.append(f"{three_digits(rem)}")

    words = " ".join(parts).strip()
    return f"{words} only/-"

def format_indian_currency(number):
    try:
        s = str(int(round(float(str(number).replace(',', '').replace('₹', '').replace('Rs.', '').replace('/-', '').strip()))))
    except Exception:
        return str(number)
    if len(s) <= 3:
        return s
    last3 = s[-3:]
    remaining = s[:-3]
    res = ""
    while len(remaining) > 2:
        res = "," + remaining[-2:] + res
        remaining = remaining[:-2]
    return remaining + res + "," + last3

print(5000000, "->", format_indian_currency(5000000), "->", indian_number_to_words(5000000))
print(1250000, "->", format_indian_currency(1250000), "->", indian_number_to_words(1250000))
print(800000, "->", format_indian_currency(800000), "->", indian_number_to_words(800000))
print(12000000, "->", format_indian_currency(12000000), "->", indian_number_to_words(12000000))


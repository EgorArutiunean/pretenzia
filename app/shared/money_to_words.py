from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP


def _plural_ru(n: int, forms: tuple[str, str, str]) -> str:
    n = abs(n) % 100
    n1 = n % 10
    if 11 <= n <= 19:
        return forms[2]
    if n1 == 1:
        return forms[0]
    if 2 <= n1 <= 4:
        return forms[1]
    return forms[2]


def _triad_to_words(n: int, gender: str = "m") -> list[str]:
    units_m = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    units_f = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    teens = [
        "десять",
        "одиннадцать",
        "двенадцать",
        "тринадцать",
        "четырнадцать",
        "пятнадцать",
        "шестнадцать",
        "семнадцать",
        "восемнадцать",
        "девятнадцать",
    ]
    tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
    hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"]

    words: list[str] = []
    h = n // 100
    t = (n % 100) // 10
    u = n % 10

    if h:
        words.append(hundreds[h])
    if t == 1:
        words.append(teens[u])
    else:
        if t:
            words.append(tens[t])
        if u:
            words.append((units_f if gender == "f" else units_m)[u])
    return words


def int_to_words_ru(n: int) -> str:
    if n == 0:
        return "ноль"
    if n < 0:
        return "минус " + int_to_words_ru(abs(n))

    groups = [
        ("", "", "", "m"),
        ("тысяча", "тысячи", "тысяч", "f"),
        ("миллион", "миллиона", "миллионов", "m"),
        ("миллиард", "миллиарда", "миллиардов", "m"),
    ]

    parts: list[str] = []
    group_index = 0
    while n > 0:
        triad = n % 1000
        if triad:
            form1, form2, form5, gender = groups[group_index]
            words = _triad_to_words(triad, gender)
            if group_index > 0:
                words.append(_plural_ru(triad, (form1, form2, form5)))
            parts = words + parts
        n //= 1000
        group_index += 1

    return " ".join(parts)


def money_to_words(value: Decimal) -> str:
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    rubles = int(value)
    kopecks = int((value - Decimal(rubles)) * 100)
    rub_word = _plural_ru(rubles, ("рубль", "рубля", "рублей"))
    kop_word = _plural_ru(kopecks, ("копейка", "копейки", "копеек"))
    return f"{int_to_words_ru(rubles)} {rub_word} {kopecks:02d} {kop_word}"

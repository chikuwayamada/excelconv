import re

# アルファベットと数字の入力チェック用正規表現
row_matcher = re.compile('^[0-9]{1,10}$')
column_matcher = re.compile('^[A-Z]{1,7}$')

# 変換に対応する最大値
VALUE_MAX = 2147483647
VALUE_MIN = 1

# アルファベットの数
N_ALPHABETS = 26


# このモジュールが発生させる例外(中身なし)
class ExcelConvError(ValueError):
    pass


# 行数をExcel列名に変換
def row_to_column(row: int):

    # 入力チェック
    if VALUE_MIN <= row <= VALUE_MAX:
        pass
    else:
        raise ExcelConvError('Out Of Range')

    result = ''

    # 入力された行数が0になるまで、最上位桁から求める。
    while row > 0:
        # 余りを計算するために文字数から1引いておく
        calc_val = row - 1

        # この桁の文字 = 余り、値の商を下位の桁へ渡す
        char_val = int(calc_val % N_ALPHABETS)
        row = int(calc_val / N_ALPHABETS)

        # Aの文字コードに余りを足して文字に変換し、列名のアルファベットにする
        result += chr(ord('A') + char_val)

    # 変換後の文字は反転しているので、並べ替えて返す
    # > strに対する[i:j:k]は、配列スライスとして動作する。
    # > [::-1]は末尾から先頭に向かって1要素ずつ文字を連結して返す動作になる
    return result[::-1]


# Excel列名を行数に変換(文字列として返す)
def column_to_row(column: str):
    # 入力チェック
    if str is None or str is '':
        raise ExcelConvError('Unknown Input')

    # 最下位の桁から計算するため、文字列を反転する
    column = column[::-1]
    result = 0
    for index in range(0, len(column)):
        char = column[index]

        # 桁の文字コード-Aの文字コード+１で10進数の数値とする
        char_value = (ord(char) - ord('A')) + 1

        if index is 0:
            # 最下位桁はそのまま加算
            result += char_value
        else:
            # 最下位以外は、値＊進数の桁数乗（**）で桁をシフトして加算する
            result += char_value * (N_ALPHABETS ** index)

    # 変換後の範囲チェック(本当は変換前にすべき
    if VALUE_MIN <= result <= VALUE_MAX:
        return str(result)
    else:
        raise ExcelConvError('Out Of Range')


# Excel列もしくは行を相互に変換
def convert_excel_axis(str_input: str):

    # 正規表現でどちらの変換をするか判定する

    # 列の判定、変換処理
    if column_matcher.match(str_input) is not None:

        return column_to_row(str_input)

    # 行の判定、変換処理
    elif row_matcher.match(str_input) is not None:

        return row_to_column(int(str_input))

    # 文字列判定に引っかからない場合は型例外
    else:
        raise ExcelConvError('Unknown Input Value Format')

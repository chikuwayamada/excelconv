import sys
import excelconv_core as exc


# usageを出力して処理を終わる
def exit_usage():
    print('Usage: # python %s [text | [option]]' % argvs[0])
    print('options')
    print('    -i : interpreter mode')
    quit()


# インタプリタモードのusageを表示
def show_interpreter_usage() :
    print('Usage: > text( 1 - 2147483647 number or A - FXSHRXW characters ) | command')
    print('commands')
    print('    /quit : quit interpreter mode')


# インタプリタモードの処理
def do_interpreter_loop() :

    while True:
        input_str = input('  >')

        # quit 以外なら変換して表示
        if input_str == 'quit' :
            print('  >bye')
            break
        else :
            # 入力値を変換して表示。例外が発生したらusage表示
            try :
                # 変換して表示
                result = exc.convert_excel_axis(input_str)
                print('  >'+str(result))
            except :
                show_interpreter_usage()
        continue
    # 標準入力を取得する

    pass

# --メイン処理の実行--
# コマンドライン引数の取得、チェック
argvs = sys.argv
arglen = len(argvs)

if arglen != 2:
    exit_usage()
elif argvs[1] == '-i':
    # インタプリタモード
    do_interpreter_loop()
else :
    # 通常変換
    try:
        result = exc.convert_excel_axis(argvs[1])
        print(result)
# 型指定なしで例外をキャッチすると、すべての例外が掛かる
#   except ValueError or TypeError:
    except :
        exit_usage()


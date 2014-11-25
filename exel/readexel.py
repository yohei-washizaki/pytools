#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd
import json
import sys
import argparse
from json import JSONEncoder

class Command:
    """Command class"""
    def __init__(self, row):
        if(row is None):
            self.command      = u'none'
            self.target       = u'none'
            self.text         = u'マルチバイト文字'
            self.animation_id = int(0)
            self.arg          = 0
        else:
            self.command      = row[0]
            self.target       = row[1]
            self.text         = row[2]
            self.animation_id = int(row[3])
            self.arg          = row[4]

    def __str__(self):
       return '{'                                 + \
              self.command.encode('utf-8') + ', ' + \
              self.target.encode('utf-8')  + ', ' + \
              self.text.encode('utf-8')    + ', ' + \
              str(self.animation_id)       + ', ' + \
              str(self.arg)                +        \
              '}';

class CommandEncoder(JSONEncoder):
    def default(self, command):
        if isinstance(command, Command):
            encoded_command = {}
            
            encoded_command['command']      = command.command
            encoded_command['target']       = command.target
            encoded_command['text']         = command.text
            encoded_command['animation-id'] = command.animation_id
            encoded_command['arg']          = command.arg
            return encoded_command
        
        else:
            return JSONEncoder.default(self, command)


def row_to_command(row):
    cmd = Command(row)
    return cmd

def main():
    # コマンドラインオプションを解析する
    parser = argparse.ArgumentParser(description='foo')
    parser.add_argument('filename', nargs='+')
    parser.add_argument('--pretty-print', action='store_true', help='pretty print output')
    parser.add_argument('--indent', nargs='?', type=int, default=4, const=4, choices=[1,2,3,4,8], help='this option is available when --pretty-print is passed')
    args = parser.parse_args('commandlist.xlsx --indent 4'.split())

    print args
    print args.filename
    print args.pretty_print
    print args.indent

    args = parser.parse_args('--help'.split())
    return

    # テストのためにファイル名は取得できたものとする
    workbook_filename = 'commandlist.xlsx'
    
    # ワークブックを開く
    workbook = xlrd.open_workbook(workbook_filename)

    # シートを取得
    sheet = workbook.sheet_by_index(0)

    # 1行ずつコマンドを解釈する
    command_list = []
    for row_index in range(0, sheet.nrows):
        row = sheet.row_values(row_index)
        cmd = row_to_command(row)
        command_list.append(cmd)

    # コマンドオブジェクトとしてまとめる
    command_set = {'__commandlist__' : command_list}

    # 出力
    text = json.dumps(command_set, cls=CommandEncoder, ensure_ascii=False, indent=2)
    print text.encode('utf-8')


main()

if __name__ == '__main__':
    main()


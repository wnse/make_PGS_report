# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:percent
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.13.7
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

# %%
import re
import pandas as pd
import logging
import argparse

pattern_chr = re.compile(r'([+-])([XY]*\d*)\(*(mos)*,*(\d+\.*\d*%)*\)*')
pattern_cnv = re.compile(r'(del|dup)\((\S+?)\)\((.*?-.*?),(.*?Mb),*(mos)*,*(\d+\.*\d*%)*\)')


# %%
def get_exp(lst):
    if len(lst) < 2:
        return [None, lst[0]]

    ty = lst[0]
    ch = lst[1]
    if ty == '+':
        if not lst[2]:
            exp = '三体'
        elif lst[2] == 'mos':
            exp = '三体嵌合'
        else:
            exp = None
    elif ty == '-':
        if not lst[2]:
            exp = '单体'
        elif lst[2] == 'mos':
            exp = '单体嵌合'
        else:
            exp = None
    elif ty == 'dup':
        if not lst[4]:
            exp = '部分重复'
        elif lst[4] == 'mos':
            exp = '部分重复嵌合'
        else:
            exp = None
    elif ty == 'del':
        if not lst[4]:
            exp = '部分缺失'
        elif lst[4] == 'mos':
            exp = '部分缺失嵌合'
        else:
            exp = None
    else:
        exp = None
    return [ch, exp]

def lst2exp(chr_num, lst, idx):
    out_dict = {}
    chr_num_tmp = 46
    if lst:
        exp_list = [get_exp(i) for i in lst]
        for ch, exp in exp_list:
            if exp:
                out_dict[exp] = out_dict.get(exp, [])
                out_dict[exp].append(ch)
            else:
                logging.error(f'illegle character:\t{lst}')
    if '三体' in out_dict.keys():
        chr_num_tmp += len(out_dict['三体'])
    if '单体' in out_dict.keys():
        chr_num_tmp -= len(out_dict['单体'])
    if chr_num:
        if chr_num != str(chr_num_tmp):
            logging.info(f'{idx} {chr_num} NOT EQUAL !!! {out_dict}')
    return out_dict


# %%
def get_schr(schr):
    note = ''
    pattern = re.compile('(\d+),(\w+)')
    if pattern.match(schr):
        (chr_num, s) = pattern.match(schr).groups()
    #if re.match('(\d+),(\w+)', schr):
        #s = re.match('(\d+),(\w+)', schr).group(1)
        # if s.upper() == 'XO':
        #     note = ' Turner综合征'
        return (chr_num, f"{s}")
    else:
        return None,None


def get_note(schr, exp_dict):
    out = ''
    note = None
    if not schr:
        note = '不推荐移植'
    elif schr and schr.upper() == 'XO':
        note = '不推荐移植'
        out = 'Turner综合征'
    elif schr and schr.upper() in ['YO']:
        note = '不推荐移植'

    total_exp_lst = []
    if exp_dict:
        for i, v in exp_dict.items():
            if v[0]:
                chrs = '、'.join(v)
                total_exp_lst.append(f'{chrs}号染色体{i}')
            else:
                total_exp_lst.append(f'{i}')

    if out and total_exp_lst:
        total_exp_lst = [out] + total_exp_lst
        out = ';'.join(total_exp_lst)
    elif total_exp_lst:
        out = ';'.join(total_exp_lst)
    elif out:
        out = out
    elif schr not in ['XY', 'XX']:
        out = ''
    else:
        out = '未见异常'

    if not note:
        if len(exp_dict) == 1:
            if re.search('嵌合', list(exp_dict.keys())[0]):
                if len(exp_dict[list(exp_dict.keys())[0]]) == 1:
                    note = '谨慎移植'
                else:
                    note = '不推荐移植'
            else:
                note = '不推荐移植'
        elif len(exp_dict) > 1:
            note = '不推荐移植'
        else:
            note = '推荐移植'
    return out, note




# %%
def dict2ext(res_dict):
    out_dict = {}
    for idx, res in res_dict.items():
        # logging.info(idx)
        out_dict[idx] = {}
        res_chr = res.strip().split(';')
        schr = res_chr.pop(0)
        chr_num, schr = get_schr(schr)
        out_dict[idx]['性染色体'] = schr
        out_dict[idx]['结果'] = res
        total_lst = []
        exp_dict = {}
        total_exp_lst = []
        note = ''
        for r in res_chr:
            if r:
                if pattern_chr.match(r):
                    total_lst.append(pattern_chr.match(r).groups())
                elif pattern_cnv.match(r):
                    total_lst.append(pattern_cnv.match(r).groups())
                else:
                    logging.error(f'{idx}\t{"not match"}')
                    total_lst.append([r])
        # if total_lst:
        exp_dict = lst2exp(chr_num, total_lst, idx)
        final_exp, note = get_note(schr, exp_dict)
        # else:

        #     for i, v in exp_dict.items():
        #         chrs = '、'.join(v)
        #         total_exp_lst.append(f'{chrs}号染色体{i}')
        # note = get_note(schr, exp_dict)
        # if total_exp_lst:
        #     tmp_exp = '; '.join(total_exp_lst) 
        # else:
        #     tmp_exp = '未见异常'
        #     note = '推荐移植'
        out_dict[idx]['解释'] = final_exp
        out_dict[idx]['备注'] = note
    return out_dict


# %%
if __name__ == '__main__':
    parse = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parse.add_argument('-i', '--input', required=True, help='intpu data including {样本} sheet which has {样本编号}&{检测结果}')
    parse.add_argument('-o', '--output', default=None, help='output excel, stdout if none')
    args = parse.parse_args()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    # f = 'test_data/test_input-0401.xlsx'
    f = args.input
    res_dict = pd.read_excel(f, '样本').set_index('样本编号')['检测结果'].to_dict()
    try:
        df_dict = dict2ext(res_dict)
    except Exception as e:
        logging.error(e)
        
    if args.output:
        try:
            pd.DataFrame.from_dict(df_dict, orient='index').to_excel(args.output)
        except Exception as e:
            logging.error(e)
    else:
        logging.info(df_dict)

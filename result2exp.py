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
    ty = lst[0]
    ch = lst[1]
    if ty == '+':
        if not lst[2]:
            exp = '重复'
        elif lst[2] == 'mos':
            exp = '三体嵌合'
        else:
            exp = None
    elif ty == '-':
        if not lst[2]:
            exp = '缺失'
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

def lst2exp(lst):
    out_dict = {}
    exp_list = [get_exp(i) for i in lst]
    for ch, exp in exp_list:
        if exp:
            out_dict[exp] = out_dict.get(exp, [])
            out_dict[exp].append(ch)
        else:
            logging.error(f'illegle character:\t{lst}')
    return out_dict


# %%
def get_schr(schr):
    if re.match('\d+,(\w+)', schr):
        s = re.match('\d+,(\w+)', schr).group(1)
        return f"{s}"
    else:
        return None


# %%
def dict2ext(res_dict):
    out_dict = {}
    for idx, res in res_dict.items():
        out_dict[idx] = {}
        res_chr = res.strip().split(';')
        schr = res_chr.pop(0)
        out_dict[idx]['性染色体'] = get_schr(schr)
        out_dict[idx]['结果'] = res
        total_lst = []
        total_exp_lst = []
        for r in res_chr:
            if r:
                if pattern_chr.match(r):
                    total_lst.append(pattern_chr.match(r).groups())
                elif pattern_cnv.match(r):
                    total_lst.append(pattern_cnv.match(r).groups())
                else:
                    logging.error(f'{idx}\t{"not match"}')
        if total_lst:
            exp_dict = lst2exp(total_lst)
            for i, v in exp_dict.items():
                chrs = '、'.join(v)
                total_exp_lst.append(f'{chrs}号染色体{i}')
        if total_exp_lst:
            tmp_exp = '; '.join(total_exp_lst) 
        else:
            tmp_exp = '未见异常'
        out_dict[idx]['解释'] = tmp_exp
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

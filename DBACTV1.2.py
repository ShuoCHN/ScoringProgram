"""
V1.2 更新：
支持对有任意评委未参评的情况进行特殊处理

"""

import xlrd

loc="inp.xlsx"

n_group=0
n_judger=0
n_viewers=0
judger_wei=1.0 # judger's mark's weight
vote_wei=1.0


marks={} #use dic to save mark <-> name (name:mark)
marks_sorted=()

def copyright():
    print("---------------------------------------------------------------------------------------------------")
    print("|                    欢迎使用【党课大活动计分程序V1.2版】 Author:ShuoCHN              ")
    print("|       本程序可实现对党课大活动评委打分及观众投票进行统计，去除评委最高分和最低分，并进行规定的加权       ")
    print("|                          最后计算出最终得分及排名，输出为txt格式                              ")

    print("|                                    使用提示：                                          ")
    print("|       1.需根据实际情况先修改inp.xlsx中的参演小组数、评委数、观众总数、评委权重、观众权重        ")
    print("|       2.将评委评分依次输入在对应组名下即可，》》将观众举旗数输入在最后一位评委下面的一行即可《《          ")
    print("|       3.使用本程序需要将inp.xlsx保存在与该exe同一目录下                                      ")
    print("|       4.完成以上工作后，按》》回车键《《即可开始进行结果的合成，结果将输出为在同一目录下的opt.txt         ")
    print("|                                                                                        ")
    print("|            在使用过程中遇到任何问题或建议请联系：计算机类205班赵硕（wxid:ShuoCHN）               ")
    print("|                             Copyright © ShuoCHN                                        ")
    print("-----------------------------------------------------------------------------------------------------\n")

def read_xls(loc):
    wb = xlrd.open_workbook(loc)
    sheet1 = wb.sheet_by_index(0)

    # get the nubmers which is important
    global n_group,n_judger,n_viewers,judger_wei,vote_wei
    n_group = int(sheet1.cell(0, 9).value)
    n_judger = int(sheet1.cell(0, 11).value)
    n_viewers = int(sheet1.cell(0, 13).value)
    judger_wei = float(sheet1.cell(0, 15).value)
    vote_wei = float(sheet1.cell(0, 17).value)

    for j in range(0,n_group):
        name = sheet1.cell(1,j).value
        judger_mark=0.0
        votes_mark = sheet1.cell(n_judger+2,j).value
        totmark=0.0
        mmax= -1
        mmin=101
        tn_judger = n_judger
        for i in range(2,n_judger+2):
            if sheet1.cell(i,j).value != "":
                mmax=max(mmax,sheet1.cell(i,j).value)
                mmin=min(mmin,sheet1.cell(i,j).value)
                judger_mark+=sheet1.cell(i,j).value
            else:
                tn_judger-=1
        judger_mark=judger_mark-mmin-mmax

        judger_mark=float(judger_mark)/float(tn_judger-2)*judger_wei

        votes_mark= (float(votes_mark)*(100.0/float(n_viewers)))*vote_wei
        totmark = judger_mark+votes_mark
        marks[name]=totmark

    # sort the marks
    global  marks_sorted
    marks_sorted = sorted(marks.items(), key=lambda kv: (kv[1], kv[0]),reverse=True)




def opt():
    txt_opt=""
    tm=marks_sorted
    for i in range(0,n_group):
        txt_temp=tm[i][0]+"最终得分："+str(format(tm[i][1], '.2f'))+",第"+str(i+1)+"名\n"
        txt_opt += txt_temp

    print(txt_opt)
    fw=open("opt.txt","w+")
    fw.write(txt_opt)


if __name__ == '__main__':

    copyright()
    input("按回车键即可开始进行结果的合成：")

    read_xls(loc)
    opt()
    input("合成已完成，请打开同目录下的opt.txt查看结果，按回车键退出")
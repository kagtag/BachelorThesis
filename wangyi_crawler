########crawler for wangyiyun
######author: Mingcheng Song
import wangyiyun.Neteasebox.api as api
from datetime import datetime
import openpyxl
import xlwt
import xlrd
import time
import random

netease=api.NetEase()


feaSeq=[]


def procHotComments(music_id):
    tmpCmt=netease.song_comments(music_id)
    tmpSeq2 = []
    for item in tmpCmt['hotComments']:

        tmpSeq = [music_id]
        tmpDict={}
        #tuple to dictionary
        for iitem in item.items():
            tmpDict[iitem[0]]=iitem[1]
        # print(tmpDict)

        if tmpDict.get('user',False):
            tmpSeq.append(tmpDict.get('user')['userId'])
        else:
            tmpSeq.append('')

        tmpSeq.append(tmpDict.get('likedCount',''))
        tmpSeq.append(datetime.fromtimestamp(tmpDict.get('time','')/1000.0).strftime('%Y-%m-%d'))
        tmpSeq.append(tmpDict.get('content',''))
        # print('success')
        tmpSeq2.append(tmpSeq)
    print(tmpSeq2)
    return tmpSeq2


        # for k in tmpDict.get('user','').items():
        #     print(k)

def getComments():
    # workbook = openpyxl.Workbook(write_only=True)
    # worksheet = workbook.create_sheet()

    workbook=xlwt.Workbook()
    worksheet=workbook.add_sheet("sheet")

    sheet=xlrd.open_workbook("wangyiyun/officialdata/songDetailFinal.xls").sheet_by_index(0)

    index=0
    for i in range(1561,sheet.nrows):
        print("第"+str(i)+"/"+str(sheet.nrows)+"首歌")
        music_id=sheet.cell_value(i,1)
        result=procHotComments(music_id)
        rsize=len(result)
        for j in range(rsize):
            for k in range(len(result[j])):
                worksheet.write(index,k,result[j][k])
            index+=1
        time.sleep(random.randint(2,3))

        if i%20==0 and i!=0:
            workbook.save("wangyiyun/tmp/comments" + str(int(i / 20 + 1)) + ".xls")
            time.sleep(random.randint(8,9))

        # if i%20==0:
        #     workbook.save("wangyiyun/tmp/comments"+str(int(i/20+1))+".xlsx")

    workbook.save("wangyiyun/tmp/comments.xls")

##将不同sheet中的数据整合到一个sheet中
def cancSheet(path,tpath):
    workbook=xlwt.Workbook()
    worksheet=workbook.add_sheet("sheet")
    book=xlrd.open_workbook(path)
    trow=0
    for i in range(book.nsheets):
        sheet=book.sheet_by_index(i)
        for j in range(sheet.nrows):
            for k in range(sheet.ncols):
                worksheet.write(trow,k,sheet.cell_value(j,k))
            trow=trow+1
    workbook.save(tpath)

def procComments(music_id):
    tmpCmt = netease.song_comments(music_id)
    for item in tmpCmt['comments']:
        print(item)


def cmtStat(path):

    workbook=xlwt.Workbook()
    worksheet=workbook.add_sheet("sheet")

    oriSheet=xlrd.open_workbook("wangyiyun/officialdata/songDetailFinal.xls").sheet_by_index(0)
    oriSongDict={}
    oriSet=set()
    for i in range(oriSheet.nrows):
        oriSongDict[oriSheet.cell_value(i,1)]=oriSheet.cell_value(i,0)
        oriSet.add(oriSheet.cell_value(i,1))
    # print(oriSongDict)
    # print()
    songDict={}
    songSet=set()
    sheet=xlrd.open_workbook(path).sheet_by_index(0)
    for i in range(sheet.nrows):
        songDict[sheet.cell_value(i,0)]=1
        songSet.add(sheet.cell_value(i,0))

    misSet=oriSet-songSet
    index=0
    for item in misSet:
        worksheet.write(index,0,item)
        worksheet.write(index,1,oriSongDict.get(item))
        # print(item, oriSongDict.get(item))
        index+=1
    # print(len(songDict.keys()))
    workbook.save("wangyiyun/officialdata/songsWithoutReviews.xls")



# procHotComments(music_id="29947420")

# print("--------------------------")
# procComments(music_id="29947420")

# getComments()
# cmtStat("wangyiyun/officialdata/commentsFinal.xls")


##########################
###########################
##############################

def user_attribute(user_id,mode=0):

    depth=0
    width=0
    user_lists=0

    while True:
        user_lists = netease.user_playlist(user_id)
        try:
            print(len(user_lists))
            break
        except:
            time.sleep(10)
            print("出现问题")
            continue

    user_dict={}
    #for each play lists
    for item in user_lists:
        # print(item['name'])

        #get information about the play list
        play_id=item['id']
        play_name=item['name']
        create_time=datetime.fromtimestamp(item['createTime']/1000.0)
        creator=[item['creator']['userId'],item['creator']['nickname']]
        creator_add=[item['creator'].get('birthday',''),item['creator'].get('gender',''),
                     item['creator'].get('province','')]
        play_count=item['playCount']
        subscribers=item['subscribers']
        subscribed_count=item['subscribedCount']

        info_seq=[play_id,play_name,subscribers,subscribed_count]
        print(info_seq)

        song_pool=[]

        try:
            # get songs'id in this play list
            for song in netease.playlist_detail(play_id):
                song_pool.append(song['id'])
        except:
            print("该歌单为空")

        user_dict[play_id]=song_pool

    return user_dict
        # for iitem in item.items():
        #     print(iitem)
        # break


def song_attribute(song_id):
    song=netease.song_detail(song_id)[0]
    cmt=netease.song_comments(song_id)
    print(cmt)

    #information
    name=song['name']
    album=[song['album']['id'],song['album']['name']]
    artist=[song['artists'][0]['id'],song['artists'][0]['name']]

    cmt_count=cmt['total']
    hotCmt=cmt['hotComments']
    normCmt=cmt['comments']

    # cmtDict={}

    for item in cmt.items():
        print(item)
    # for item in song.items():
    #     print(item)

def comment_attribute(cmt_dict):
    tmpSeq=[]

    tmpSeq.append(cmt_dict['user'].get('userId',''))
    tmpSeq.append(cmt_dict.get('likedCount', ''))
    tmpSeq.append(datetime.fromtimestamp(cmt_dict.get('time', '') / 1000.0).strftime('%Y-%m-%d'))
    tmpSeq.append(cmt_dict.get('content', ''))
    return tmpSeq

#compute depth and width of a playlist
def playlist_attribute(playlist_id):
    play=netease.playlist_detail(playlist_id)
    print(len(play))
    for item in play:
        for iitem in item.items():
            print(iitem)
        break


def generate_song_set():
    workbook = openpyxl.Workbook()
    worksheet = workbook.create_sheet()

    sheet = xlrd.open_workbook("wangyiyun/officialdata/songDetailFinal.xls").sheet_by_index(0)

    index = 1
    for i in range(176, sheet.nrows):
        print("第" + str(i) + "/" + str(sheet.nrows) + "首歌---------------------------------")
        worksheet.cell(row=index, column=1).value =str(i)
        index+=1
        music_id = sheet.cell_value(i, 1)
        tmpCmt = netease.song_comments(music_id)

        try:
            hot_cmt = tmpCmt['hotComments']
            norm_cmt = tmpCmt['comments']
        except:
            continue
        print("共有" + str(len(hot_cmt)) + "位用户发表了热评")
        for item in hot_cmt:
            if hot_cmt=='':
                continue

            print("--------------------------------")
            tmp_user_id=item['user']['userId']

            # print('用户'+str(tmp_user_id))
            user_dict=user_attribute(tmp_user_id)
            print(str(tmp_user_id) + ' 用户歌单数量 ' + str(len(user_dict)))
            print("--------------------------------")
            time.sleep(random.randint(2,3))#rest for 2 sec for each user
            for lis in user_dict.items():
                # if index>3:
                #     break

                worksheet.cell(row=index,column=1).value=tmp_user_id
                worksheet.cell(row=index,column=2).value=lis[0]
                worksheet.cell(row=index,column=3).value=str(lis[1])
                index+=1

                if index%50==0 and index!=0:
                    workbook.save("wangyiyun/tmp/user_data"+str(int(index/50+1))+".xlsx")

                # for song in lis:
                #     print(tmp_user_id,lis,song)
            # print(user_dict)
            # print("success final")
        #     break
        # break
    workbook.save('wangyiyun/tmp/user_data.xlsx')



generate_song_set()
# test_user_id="48548007"
# print(user_attribute(test_user_id))
# print("--------------------------------")
# def main():



# playlist_attribute("44239099")

# song_attribute("27646205")


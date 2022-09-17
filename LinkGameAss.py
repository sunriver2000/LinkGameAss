# coding=utf-8

from pymouse import PyMouse
from PIL import ImageGrab, Image
from win32.lib import win32con

import win32gui, win32print
import os, time
import operator

import numpy as np
import sys
import argparse

def save_im(image, name):
    # 创建存储路径
    screen_path = os.path.join(os.path.dirname(__file__), 'screen')
    if not os.path.exists(screen_path):
        os.makedirs(screen_path)
    
    # 保存图片到存储路径
    image_name = os.path.join(screen_path, name)
    #image_name = screen_path

    #print("image_name: %", image_name)
    t = time.strftime('%Y%m%d_%H%M%S', time.localtime())
    image.save('%s_%s.png' % (image_name, t))  # 文件名name后面加了个时间戳，避免重名

class Game:
    # 初始化
    def __init__(self):
        self.im2num_arr = np.zeros((10, 14), dtype=np.int32)
        self.im2num_arr_backup = np.zeros((10, 14), dtype=np.int32)
        self.imTypeInGame = [12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 34, 38, 42]
        self.isGetStep = False
        self.candidateStepNum = 0
        self.candidateStepArr = np.zeros((8, 5), dtype=np.int32)

    # 计算多边形周长
    def calc(self):
        z = 0
        for i in range(1, 9):
            for j in range(1, 13):
                if self.im2num_arr[i][j] != 0:
                    if self.im2num_arr[i-1][j] == 0:
                        z += 1
                    if self.im2num_arr[i+1][j] == 0:
                        z += 1
                    if self.im2num_arr[i][j-1] == 0:
                        z += 1
                    if self.im2num_arr[i][j+1] == 0:
                        z += 1
        return z

    # 统计解法数量
    def calcStepNum(self):
        xx = [4, 5, 3, 6, 2, 7, 1, 8]
        yy = [6, 7, 5, 8, 4, 9, 3, 10, 2, 11, 1, 12]
        
        z = 0
        for x1 in xx:
            for y1 in yy:
                if self.im2num_arr[x1][y1] == 0:
                    continue

                for x2 in xx:
                    for y2 in yy:
                        # 跳过为0 或者同一个
                        if self.im2num_arr[x2][y2] == 0 or (x1 == x2 and y1 == y2):
                            continue

                        if self.isReachable(x1, y1, x2, y2):  
                            z += 1
        return z  
 
    # 筛选
    def getBestStep(self):
        for i in range(0, 10):
            for j in range(0, 14):
                self.im2num_arr_backup[i][j] = self.im2num_arr[i][j]
    
        a = 0
        b = 0
        best = 0 
        while a < self.candidateStepNum:
            self.run(self.candidateStepArr[a][0], self.candidateStepArr[a][1], self.candidateStepArr[a][2], self.candidateStepArr[a][3], level)
            #self.candidateStepArr[a][4] = self.calcStepNum()
            self.candidateStepArr[a][4] = self.calc()
            
            if self.candidateStepArr[a][4] > best:
                best = self.candidateStepArr[a][4]
                b = a

            for i in range(0, 10):
                for j in range(0, 14):
                    self.im2num_arr[i][j] = self.im2num_arr_backup[i][j]
    
            a += 1
       
        print(self.candidateStepArr)

        return self.candidateStepArr[b][0], self.candidateStepArr[b][1], self.candidateStepArr[b][2], self.candidateStepArr[b][3]
            
    # 找下一步解法
    def findNextStep(self, level):
        self.findCandidateStepArr(level)
        return self.getBestStep()

    # 是否存在
    def isExist(self, x1, y1, x2, y2):
        a = 0
           
        while a < self.candidateStepNum:
            if x1 == self.candidateStepArr[a][2] and y1 == self.candidateStepArr[a][3] and x2 == self.candidateStepArr[a][0] and y2 == self.candidateStepArr[a][1]:
                return True

            a += 1
        return False    

    # 找候选解法
    def findCandidateStepArr(self, level):
        xx = [4, 5, 3, 6, 2, 7, 1, 8]
        yy = [6, 7, 5, 8, 4, 9, 3, 10, 2, 11, 1, 12]
        
        self.isGetStep = False
        self.candidateStepNum = 0
        self.candidateStepArr = np.zeros((8, 5), dtype=np.int32)

        for x1 in xx:
            for y1 in yy:
                if self.im2num_arr[x1][y1] == 0:
                    continue

                for x2 in xx:
                    for y2 in yy:
                        # 跳过为0 或者同一个
                        if self.im2num_arr[x2][y2] == 0 or (x1 == x2 and y1 == y2):
                            continue

                        if self.isReachable(x1, y1, x2, y2):
                            self.isGetStep = True

                            if self.isExist(x1, y1, x2, y2):
                                continue

                            self.candidateStepArr[self.candidateStepNum][0] = x1
                            self.candidateStepArr[self.candidateStepNum][1] = y1
                            self.candidateStepArr[self.candidateStepNum][2] = x2
                            self.candidateStepArr[self.candidateStepNum][3] = y2
                            self.candidateStepArr[self.candidateStepNum][4] = 0
                            
                            self.candidateStepNum = self.candidateStepNum + 1
                            if self.candidateStepNum == 8:
                                return
                                
    # 判断图标类型数量是否正确
    def imTypeIsRight(self, level, imType):
        print("level = {}, imType = {}".format(level, imType))
        if self.imTypeInGame[level-1] == imType:
            return True
        return False    

    # 游戏内核运行
    def run(self, x1, y1, x2, y2, level):
        self.im2num_arr[x1][y1] = 0
        self.im2num_arr[x2][y2] = 0

        if level == 11 or level == 13:
            runtimes = 2
        else:
            runtimes = 1

        rt = 0
        while rt < runtimes:        
            for a in range(0, 2):
                for x in range(1, 9):
                    for y in range(1, 13):
                        if self.im2num_arr[x][y] == 0:
                            if level == 2:
                                self.level02(x, y)
                            elif level == 3:
                                self.level03(x, y)
                            elif level == 4:
                                self.level04(x, y)
                            elif level == 5:
                                self.level05(x, y)    
                            elif level == 6:
                                self.level06(x, y)
                            elif level == 7:
                                self.level07(x, y) 
                            elif level == 8:
                                self.level08(x, y)
                            elif level == 9:
                                self.level09(x, y)
                            elif level == 10:
                                self.level10(x, y)   
                            elif level == 11:
                                if rt == 0:
                                    self.level06(x, y)
                                else:
                                    self.level08(x, y)
                            elif level == 12:
                                self.level12(x, y)
                            elif level == 13:
                                if rt == 0:
                                    self.level07(x, y)
                                else:
                                    self.level09(x, y)
            rt += 1                                     

    # 第二关 向下
    def level02(self, x, y):
        for z in range(x, 0, -1):
            self.im2num_arr[z][y] = self.im2num_arr[z-1][y]

    # 第三关 向左
    def level03(self, x, y):
        for z in range(y, 13):
            self.im2num_arr[x][z] = self.im2num_arr[x][z+1]

    # 第四关 向上
    def level04(self, x, y):
        for z in range(x, 9):
            self.im2num_arr[z][y] = self.im2num_arr[z+1][y]    

    # 第五关 向右
    def level05(self, x, y):
        for z in range(y, 0, -1):
            self.im2num_arr[x][z] = self.im2num_arr[x][z-1] 

    # 第六关 上下分离
    def level06(self, x, y):
        if x < 5:
            for z in range(x, 4):
                self.im2num_arr[z][y] = self.im2num_arr[z+1][y]
            self.im2num_arr[4][y] = 0
        else:
            for z in range(x, 5, -1):
                self.im2num_arr[z][y] = self.im2num_arr[z-1][y]
            self.im2num_arr[5][y] = 0    

    # 第七关 上下靠拢
    def level07(self, x, y):
        if x < 5:
            for z in range(x, 0, -1):
                self.im2num_arr[z][y] = self.im2num_arr[z-1][y]
        else:
            for z in range(x, 9):
                self.im2num_arr[z][y] = self.im2num_arr[z+1][y]          

    # 第八关 左右分离
    def level08(self, x, y):
        if y < 7:
            for z in range(y, 6):
                self.im2num_arr[x][z] = self.im2num_arr[x][z+1]
            self.im2num_arr[x][6] = 0
        else:
            for z in range(y, 7, -1):
                self.im2num_arr[x][z] = self.im2num_arr[x][z-1]
            self.im2num_arr[x][7] = 0         

    # 第九关 左右靠拢
    def level09(self, x, y):
        if y < 7:
            for z in range(y, 0, -1):
                self.im2num_arr[x][z] = self.im2num_arr[x][z-1]
        else:
            for z in range(y, 13):
                self.im2num_arr[x][z] = self.im2num_arr[x][z+1]

    # 第十关 横向错位
    def level10(self, x, y):
        if x < 5:
            for z in range(y, 13):
                self.im2num_arr[x][z] = self.im2num_arr[x][z+1]
        else:
            for z in range(y, 0, -1):
                self.im2num_arr[x][z] = self.im2num_arr[x][z-1] 
    
    # 第十二关 纵向错位
    def level12(self, x, y):
        if y < 7:
            for z in range(x, 9):
                self.im2num_arr[z][y] = self.im2num_arr[z+1][y]
        else:
            for z in range(x, 0, -1):
                self.im2num_arr[z][y] = self.im2num_arr[z-1][y] 


    # 是否为同行或同列且相连
    def isReachable(self, x1, y1, x2, y2):
        
        #1、先判断值是否相同
        if self.im2num_arr[x1][y1] != self.im2num_arr[x2][y2]:
            return False

        #2、相邻的或者同行列的
        if self.isDirectConnect(x1, y1, x2, y2):
            return True

        # 3、分别获取两个坐标同行或同列可连的坐标数组
        list1 = self.getDirectConnectList(x1, y1)
        list2 = self.getDirectConnectList(x2, y2)

        # 4、比较坐标数组中是否可连
        for x1, y1 in list1:
            for x2, y2 in list2:
                if self.isDirectConnect(x1, y1, x2, y2):
                    return True

        return False

    # 获取同行或同列可连的坐标数组
    def getDirectConnectList(self, x, y):
        plist = []

        for px in range(0, 10):
            for py in range(0, 14):
                # 获取同行或同列且为0的坐标
                if self.im2num_arr[px][py] == 0 and self.isDirectConnect(x, y, px, py):
                    plist.append([px, py])
        
        return plist

    # 是否为同行或同列且可连
    def isDirectConnect(self, x1, y1, x2, y2):
        # 1、位置完全相同
        if x1 == x2 and y1 == y2:
            return False

        # 2、行列都不同的
        if x1 != x2 and y1 != y2:
            return False

        # 3、同行
        if x1 == x2 and self.isRowConnect(x1, y1, y2):
            return True

        # 4、同列
        if y1 == y2 and self.isColConnect(y1, x1, x2):
            return True

        return False

    # 判断同行是否可连
    def isRowConnect(self, x, y1, y2):
        minY = min(y1, y2)
        maxY = max(y1, y2)

        # 相邻直接可连
        if maxY - minY == 1:
            return True

        # 判断两个坐标之间是否全为0
        for y0 in range(minY + 1, maxY):
            if self.im2num_arr[x][y0] != 0:
                return False
        
        return True

    # 判断同列是否可连
    def isColConnect(self, y, x1, x2):
        minX = min(x1, x2)
        maxX = max(x1, x2)

        # 相邻直接可连
        if maxX - minX == 1:
            return True

        # 判断两个坐标之间是否全为0
        for x0 in range(minX + 1, maxX):
            if self.im2num_arr[x0][y] != 0:
                return False
 
        return True

    # 判断矩阵是否全为0
    def isAllZero(self):
        z = 0
        for i in range(1, 9):
            for j in range(1, 13):
                if self.im2num_arr[i][j] != 0:
                    z = z + 1
                if z > 1:    
                    return False

        return True

class GameAssist:
    #初始化
    def __init__(self, wdname):
        
        # 取得窗口句柄
        self.hwnd = win32gui.FindWindow(None, wdname)

        if not self.hwnd:
            print("can't find window, name is : 【%s】" % wdname)
            exit()

        print("find window, name is : 【%s】" % wdname)

        # 窗口显示最前面
        #win32gui.SetForegroundWindow(self.hwnd)

        # 小图标样本
        self.im_type_list = []
        self.im2num_type_lst = []
        self.imType = 0
        self.imageNum = 0

        # 小图标宽高
        self.im_width = 0

        self.result = 0
        self.hash1 = 0
        self.hash2 = 0

        # 左上角坐标
        self.x0 = 0
        self.y0 = 0

        # 右下角坐标
        self.x1 = 0
        self.y1 = 0 

        self.mouse = PyMouse()

        # 获取分辨率
        hDC = win32gui.GetDC(0)
        self.w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
        self.h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
        
        print("w = {}, h = {}".format(self.w, self.h))

    # 锁定图标矩阵
    def lockImage(self):
        image = ImageGrab.grab((0, 0, self.w-1, self.h-1))
        image = image.resize((self.w, self.h), Image.ANTIALIAS).convert("L")

        self.x0 = 0
        self.y0 = 0
        self.x1 = 0
        self.y1 = 0

        threshold = 250

        table = []
        for i in range(256):
            if i < threshold:
                table.append(0)
            else:
                table.append(1)

        image1 = image.point(table, '1')

        # 获取左上角坐标
        for j in range(200, 300):
            for i in range(300, 380):
        
                if image1.getpixel((i,j)) == 1:
                    self.x0 = i - 2
                    self.y0 = j - 2
                    print("x = {}, y = {}".format(i, j))
                    break
            else:
                continue
            break

        # 获取右下角坐标
        for j in range(900, 800, -1):
            for i in range(1300, 1200, -1):
            
                if image1.getpixel((i,j)) == 1:
                    self.x1 = i + 2
                    self.y1 = j + 2
                    print("x = {}, y = {}".format(i, j))
                    break
            else:
                continue
            break        

        # 计算图标边长，高度=宽度
        self.im_width = (self.x1 - self.x0) / 12

        print("im_width = {}".format(self.im_width))

        #save_im(image1, 'all')

        if self.x0 == 0 or self.y1 == 0 or self.x1 == 0 or self.y1 == 0:
            return False

        return True


    # 分割图标
    def screenshot(self):
        image = ImageGrab.grab((self.x0, self.y0, self.x1, self.y1))
        #save_im(image, 'image')

        image_list = {}
        offset = self.im_width

        for x in range(8):
            image_list[x] = {}
            for y in range(12):
                top = round(x * offset) + 3
                left = round(y * offset) + 3
                bottom = round((x + 1) * offset) - 3
                right = round((y + 1) * offset) - 3

                im = image.crop((left, top, right, bottom))
                image_list[x][y] = im

                #image1 = image_list[x][y].resize((20, 20), Image.ANTIALIAS).convert("L")
                #save_im(image1, str(x).zfill(2)+str(y).zfill(2))

                #print("top = {}, left = {}".format(top, left))
                #save_im(image_list[x][y], str(x).zfill(2)+str(y).zfill(2))
                
        return image_list

    # 获取图标库
    def getImageType(self, game, image_list):
        self.im_type_list = []
        for i in range(len(image_list)):
            for j in range(len(image_list[0])):
                index = self.getIndex(image_list[i][j])
                if index < 0:
                    self.im_type_list.append(image_list[i][j])

        return len(self.im_type_list)
                    

    # 标识矩阵
    def image2num(self, game, image_list):
        imageNum = 0
        # 1、识别出不同的图片，将图片矩阵转换成数字矩阵
        for i in range(len(image_list)):
            for j in range(len(image_list[0])):
                index = self.getIndex(image_list[i][j])

                if index < 0:
                    game.im2num_arr[i+1][j+1] = 0
                else:
                    game.im2num_arr[i+1][j+1] = index + 1
                    imageNum = imageNum + 1   


        #for i in range(len(self.im_type_list)):
        #    save_im(self.im_type_list[i], str(i).zfill(2))

        #self.analysis(game, image_list)


        print("当前图标数：{}".format(imageNum))        
        print(game.im2num_arr)

        return imageNum
        
    # 分析图像识别算法
    def analysis(self, game, image_list):
        
        image = ImageGrab.grab((self.x0, self.y0, self.x1, self.y1))
        print("{},{},{},{}".format(self.x0, self.y0, self.x1, self.y1))
        save_im(image, 'image')

        #for x in range(8):
        #    for y in range(12):
        #        save_im(image_list[x][y], str(x).zfill(2)+str(y).zfill(2))

        #for i in range(len(self.im_type_list)):
        #    save_im(self.im_type_list[i], str(i).zfill(2))

        f = open('log.txt','w')
        for z in range(len(self.im_type_list)):

            for i in range(len(image_list)):
                for j in range(len(image_list[0])):
                    if game.im2num_arr[i+1][j+1] == z + 1:
                        
                        self.isMatch(self.im_type_list[z], image_list[i][j])
                        f.write(str(z + 1) + ' ' + str(self.result).zfill(3) + ' ' + str(self.hash2) + '\n')
            f.write('\n')
        f.close()  


    # 检查数组中是否有图标，如果有则返回索引
    def getIndex(self, im):
        for i in range(len(self.im_type_list)):
            if self.isMatch(im, self.im_type_list[i]):
                return i

        return -1

    # 汉明距离判断两个图标是否一样
    def isMatch(self, im1, im2):
        # 缩小图标，转成灰度
        image1 = im1.resize((20, 20), Image.ANTIALIAS).convert("L")
        image2 = im2.resize((20, 20), Image.ANTIALIAS).convert("L")

        # 降灰度图标转成01串，即二进制数据
        pixels1 = list(image1.getdata())
        pixels2 = list(image2.getdata())

        avg1 = sum(pixels1) / len(pixels1)    
        avg2 = sum(pixels2) / len(pixels2)    
        self.hash1 = "".join(map(lambda p: "1" if p > avg1 else "0", pixels1))
        self.hash2 = "".join(map(lambda p: "1" if p > avg2 else "0", pixels2))

        # 统计两个01串不同数据的个数
        match = sum(map(operator.ne, self.hash1, self.hash2))

        self.result = match

        # 阀值设为10
        return match < 40

    # 点击事件并设置数组为0
    def clickAndSetZero(self, x1, y1, x2, y2):
        # print("click", x1, y1, x2, y2)

        # 原理：左上角图标中点 + 偏移量
        p1_x = int(self.x0 + (y1 - 1)*self.im_width + (self.im_width / 2))
        p1_y = int(self.y0 + (x1 - 1)*self.im_width + (self.im_width / 2))

        p2_x = int(self.x0 + (y2 - 1)*self.im_width + (self.im_width / 2))
        p2_y = int(self.y0 + (x2 - 1)*self.im_width + (self.im_width / 2))
        
        time.sleep(0.2)
        self.mouse.click(p1_x, p1_y)
        time.sleep(0.2)
        self.mouse.click(p2_x, p2_y)
        
        print("消除：(%02d, %02d) (%02d, %02d)" % (x1, y1, x2, y2))
        # exit()

    def resetMouse(self):
        p3_x = int(self.x0 + (0 - 1)*self.im_width + (self.im_width / 2))
        p3_y = int(self.y0 + (0 - 1)*self.im_width + (self.im_width / 2))
        
        time.sleep(0.2)
        self.mouse.click(p3_x, p3_y)
        time.sleep(0.5)


    # 程序入口、控制中心
    def start(self, game, level, mode):

        # 1、锁定图标矩阵
        if not self.lockImage():
            return False

        # 2、先截取游戏区域大图，然后分切每个小图
        image_list = self.screenshot()

        # 3、获取图标库
        self.imType = self.getImageType(game, image_list)
        print("图标数：", self.imType)

        # 4、识别小图标，收集编号
        self.imageNum = self.image2num(game, image_list)

        if game.imTypeIsRight(level, self.imType):
            print("start to play.")
        else:
            #self.analysis(game, image_list)
            print("error, end.")
            return False
        
        step = 0 
        
        # 4、遍历查找可以相连的坐标
        while not game.isAllZero():
            x1, y1, x2, y2 = game.findNextStep(level)
            if game.isGetStep:
                self.clickAndSetZero(x1, y1, x2, y2)
                step = step + 2
                print(str(step))
                
                if mode == 'a':
                    game.run(x1, y1, x2, y2, level)
                    print(game.im2num_arr)
                    
                    # 检查
                    if step % 16 == 0:
                        print("check")
                        self.resetMouse()

                        image_list = self.screenshot()
                        self.imageNum = self.image2num(game, image_list)
                        
                        step = 96 - self.imageNum
                elif mode == 'm':
                    time.sleep(0.6)
                    image_list = self.screenshot()
                    self.imageNum = self.image2num(game, image_list)
                    
            else:
                # 洗牌
                print("no step")
                self.resetMouse()

                image_list = self.screenshot()
                self.imageNum = self.image2num(game, image_list)
                step = 96 - self.imageNum

        print("end")
        return True


if __name__ == "__main__":
    parse = argparse.ArgumentParser(prog='模式')
    parse.add_argument('-m', dest = 'mode', type = str, help='模式')

    result = parse.parse_args()
    mode = result.mode

    print(str(mode))
    if mode != 'm' and mode != 'a': 
        sys.exit()

    wdname = u'宠物连连看经典版2小游戏,在线玩,4399小游戏 - 360极速浏览器 13.5'
    
    game = Game()
    demo = GameAssist(wdname)
    
    #cmd = str(input("请输入命令："))

    level = 1 

    while 1:
        #if level == 14 or cmd == 'q':
        if level == 14:
            print("game over.")
            sys.exit()
 
        print("level = ", level)

        if not demo.start(game, level, mode):
            print("game exit.")
            sys.exit()

        time.sleep(5)
        demo.mouse.click(820, 820)
        level = level + 1
        time.sleep(2)

        #cmd = str(input("请输入命令："))

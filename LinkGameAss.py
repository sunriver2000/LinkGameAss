# coding=utf-8

from pymouse import PyMouse
from PIL import ImageGrab, Image
from win32.lib import win32con

import win32gui, win32print
import os, time
import operator

import numpy as np
import sys

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
        self.im2num_arr = []
        self.imTypeInGame = [12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 34, 38, 42]
        self.isGetStep = False

    # 找下一步解法
    def findNextStep(self):
        xx = [4, 5, 3, 6, 2, 7, 1, 8]
        yy = [6, 7, 5, 8, 4, 9, 3, 10, 2, 11, 1, 12]
        self.isGetStep = False
    
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
                            return x1, y1, x2, y2

        return x1, y1, x2, y2

    # 判断图标类型数量是否正确
    def imTypeIsRight(self, level, imType):
        print("level = {}, imType = {}".format(level, imType))
        if self.imTypeInGame[level-1] == imType:
            return True
        return False    

    # 游戏内核运行
    def run(self, x1, y1, x2, y2, level):
        if level == 1:
            self.level1(x1, y1, x2, y2)
        elif level == 2:
            self.level2(x1, y1, x2, y2)
        elif level == 3:
            self.level3(x1, y1, x2, y2)
        elif level == 4:
            self.level4(x1, y1, x2, y2)
        elif level == 5:
            self.level5(x1, y1, x2, y2)    
        elif level == 6:
            self.level6(x1, y1, x2, y2)
        elif level == 7:
            self.level7(x1, y1, x2, y2) 
        elif level == 8:
            self.level8(x1, y1, x2, y2)
        elif level == 9:
            self.level9(x1, y1, x2, y2)
        elif level == 10:
            self.level10(x1, y1, x2, y2)                    

    # 第一关
    def level1(self, x1, y1, x2, y2):
        # 设置矩阵值为0
        self.im2num_arr[x1][y1] = 0
        self.im2num_arr[x2][y2] = 0

    # 第二关
    def level2(self, x1, y1, x2, y2):
        for x in range(x1, 0, -1):
            self.im2num_arr[x][y1] = self.im2num_arr[x-1][y1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if x - 1 == x2 and y1 == y2:
                x2 = x

        for x in range(x2, 0, -1):
            self.im2num_arr[x][y2] = self.im2num_arr[x-1][y2]       
    
    # 第三关
    def level3(self, x1, y1, x2, y2):
        for y in range(y1, 13):
            self.im2num_arr[x1][y] = self.im2num_arr[x1][y+1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if y + 1 == y2 and x1 == x2:
                y2 = y 

        for y in range(y2, 13):
            self.im2num_arr[x2][y] = self.im2num_arr[x2][y+1]   

    # 第四关
    def level4(self, x1, y1, x2, y2):
        for x in range(x1, 9):
            self.im2num_arr[x][y1] = self.im2num_arr[x+1][y1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if x + 1 == x2 and y1 == y2:
                x2 = x

        for x in range(x2, 9):
            self.im2num_arr[x][y2] = self.im2num_arr[x+1][y2]     

    # 第五关
    def level5(self, x1, y1, x2, y2):
        for y in range(y1, 0, -1):
            self.im2num_arr[x1][y] = self.im2num_arr[x1][y-1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if y - 1 == y2 and x1 == x2:
                y2 = y 

        for y in range(y2, 0, -1):
            self.im2num_arr[x2][y] = self.im2num_arr[x2][y-1]

    # 第六关
    def level6(self, x1, y1, x2, y2):
        for x in range(x1, 4):
            self.im2num_arr[x][y1] = self.im2num_arr[x+1][y1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if x + 1 == x2 and y1 == y2:
                x2 = x

        if x1 < 5:
            self.im2num_arr[4][y1] = 0                      

        for x in range(x2, 4):
            self.im2num_arr[x][y2] = self.im2num_arr[x+1][y2]

        if x2 < 5:
            self.im2num_arr[4][y2] = 0

        ###
        for x in range(x1, 5, -1):
            self.im2num_arr[x][y1] = self.im2num_arr[x-1][y1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if x - 1 == x2 and y1 == y2:
                x2 = x
        if x1 > 4:
            self.im2num_arr[5][y1] = 0

        for x in range(x2, 5, -1):
            self.im2num_arr[x][y2] = self.im2num_arr[x-1][y2]

        if x2 > 4:
            self.im2num_arr[5][y2] = 0

    # 第七关
    def level7(self, x1, y1, x2, y2):
        if x1 < 5:
            for x in range(x1, 0, -1):
                self.im2num_arr[x][y1] = self.im2num_arr[x-1][y1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if x - 1 == x2 and y1 == y2:
                    x2 = x
        if x2 < 5:
            for x in range(x2, 0, -1):
                self.im2num_arr[x][y2] = self.im2num_arr[x-1][y2] 
        
        if x1 > 4:
            for x in range(x1, 9):
                self.im2num_arr[x][y1] = self.im2num_arr[x+1][y1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if x + 1 == x2 and y1 == y2:
                    x2 = x
        if x2 > 4:
            for x in range(x2, 9):
                self.im2num_arr[x][y2] = self.im2num_arr[x+1][y2]          

    # 第八关
    def level8(self, x1, y1, x2, y2):
        for y in range(y1, 6):
            self.im2num_arr[x1][y] = self.im2num_arr[x1][y+1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if y + 1 == y2 and x1 == x2:
                y2 = y

        if y1 < 7:
            self.im2num_arr[x1][6] = 0          

        for y in range(y2, 6):
            self.im2num_arr[x2][y] = self.im2num_arr[x2][y+1]

        if y2 < 7:
            self.im2num_arr[x2][6] = 0    

        ###
        for y in range(y1, 7, -1):
            self.im2num_arr[x1][y] = self.im2num_arr[x1][y-1]

            # 判断第二个气泡是否移动，如果移动则坐标更新
            if y - 1 == y2 and x1 == x2:
                y2 = y 

        if y1 > 6:
            self.im2num_arr[x1][7] = 0

        for y in range(y2, 7, -1):
            self.im2num_arr[x2][y] = self.im2num_arr[x2][y-1]

        if y2 > 6:
            self.im2num_arr[x2][7] = 0

    # 第九关
    def level9(self, x1, y1, x2, y2):
        if y1 < 7:
            for y in range(y1, 0, -1):
                self.im2num_arr[x1][y] = self.im2num_arr[x1][y-1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if y - 1 == y2 and x1 == x2:
                    y2 = y
        if y2 < 7:
            for y in range(y2, 0, -1):
                self.im2num_arr[x2][y] = self.im2num_arr[x2][y-1] 
        
        if y1 > 6:
            for y in range(y1, 13):
                self.im2num_arr[x1][y] = self.im2num_arr[x1][y+1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if y + 1 == y2 and x1 == x2:
                    y2 = y
        if y2 > 6:
            for y in range(y2, 13):
                self.im2num_arr[x2][y] = self.im2num_arr[x2][y+1]    

    # 第十关
    def level10(self, x1, y1, x2, y2):
        if x1 < 5:
            for y in range(y1, 13):
                self.im2num_arr[x1][y] = self.im2num_arr[x1][y+1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if y + 1 == y2 and x1 == x2:
                    y2 = y 
        if x2 < 5:
            for y in range(y2, 13):
                self.im2num_arr[x2][y] = self.im2num_arr[x2][y+1]  
                 
        if x1 > 4:
            for y in range(y1, 0, -1):
                self.im2num_arr[x1][y] = self.im2num_arr[x1][y-1]

                # 判断第二个气泡是否移动，如果移动则坐标更新
                if y - 1 == y2 and x1 == x2:
                    y2 = y 
        if x2 > 4:
            for y in range(y2, 0, -1):
                self.im2num_arr[x2][y] = self.im2num_arr[x2][y-1] 

                        
    
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
        for i in range(1, 9):
            for j in range(1, 13):
                if self.im2num_arr[i][j] != 0:
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

    # 锁定图标矩阵
    def lockImage(self):
        # 获取分辨率
        hDC = win32gui.GetDC(0)
        w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
        h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
        
        print("w = {}, h = {}".format(w, h))

        image = ImageGrab.grab((0, 0, w-1, h-1))

        image = image.resize((w, h), Image.ANTIALIAS).convert("L")

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

        save_im(image1, 'all')

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

                image1 = image_list[x][y].resize((20, 20), Image.ANTIALIAS).convert("L")
                
                #save_im(image1, str(x).zfill(2)+str(y).zfill(2))

                #print("top = {}, left = {}".format(top, left))
                #save_im(image_list[x][y], str(x).zfill(2)+str(y).zfill(2))
                
        return image_list

    # 标识矩阵
    def image2num(self, game, image_list):
        # 1、创建全零矩阵和空的一维数组
        arr = np.zeros((10, 14), dtype=np.int32)
        self.im_type_list = []

        # 2、识别出不同的图片，将图片矩阵转换成数字矩阵
        for i in range(len(image_list)):
            for j in range(len(image_list[0])):
                im = image_list[i][j]

                # 验证当前图标是否已存入
                index = self.getIndex(im)

                # 不存在，则存入
                if index < 0:
                    self.im_type_list.append(im)
                    arr[i+1][j+1] = len(self.im_type_list)
                else:
                    arr[i+1][j+1] = index + 1

        print("图标数：", len(self.im_type_list))

        #for i in range(len(self.im_type_list)):
        #    save_im(self.im_type_list[i], str(i).zfill(2))

        game.im2num_arr = arr
        print(game.im2num_arr)

        #self.analysis(game, image_list)

        self.imType = len(self.im_type_list)


    # 洗牌后重新识别
    def image2num2(self, game, image_list):
        imNum = 0
        for i in range(len(image_list)):
            for j in range(len(image_list[0])):
                im = image_list[i][j]

                # 验证当前图标是否已存入
                index = self.getIndex(im)
                if index < 0:
                    game.im2num_arr[i+1][j+1] = 0
                else:
                    game.im2num_arr[i+1][j+1] = index + 1
                    imNum = imNum + 1                    

        print(game.im2num_arr)
        return imNum

    def image2num3(self, game, image_list):
        arr1 = np.zeros((10, 14), dtype=np.int32)
        
        imNum = 0
        for i in range(len(image_list)):
            for j in range(len(image_list[0])):
                im = image_list[i][j]

                # 验证当前图标是否已存入
                index = self.getIndex(im)
                if index < 0:
                    arr1[i+1][j+1] = 0
                else:
                    arr1[i+1][j+1] = index + 1
                    imNum = imNum + 1                    

        print(arr1)
        return imNum
        
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

    # 程序入口、控制中心
    def start(self, game, level):

        # 1、锁定图标矩阵
        if not self.lockImage():
            return False

        # 2、先截取游戏区域大图，然后分切每个小图
        image_list = self.screenshot()

        # 3、识别小图标，收集编号
        self.image2num(game, image_list)

        if game.imTypeIsRight(level, self.imType):
            print("start to play.")
        else:
            self.analysis(game, image_list)
            print("error, end.")
            return False
        
        step = 0 
        
        # 4、遍历查找可以相连的坐标
        while not game.isAllZero():
            x1, y1, x2, y2 = game.findNextStep()
            if game.isGetStep:
                self.clickAndSetZero(x1, y1, x2, y2)
                step = step + 1
                
                if level < 11:
                    game.run(x1, y1, x2, y2, level)
                else:
                    time.sleep(1)
                    image_list = self.screenshot()
                    self.image2num2(game, image_list)

                    '''
                if step % 8 == 0:
                    time.sleep(1)
                    image_list = self.screenshot()
                    print(game.im2num_arr)
                    print("-------------------------------")
                    imNum = self.image2num3(game, image_list)
                    if step * 2 + imNum != 96:
                        print("step = {}, imNum = {}".format(step, imNum))
                        return False
'''


            else:
                print("no step")
                time.sleep(2)

                image_list = self.screenshot()
                self.image2num2(game, image_list)

        print("end")
        return True


if __name__ == "__main__":
    wdname = u'宠物连连看经典版2小游戏,在线玩,4399小游戏 - 360极速浏览器 13.5'
    
    game = Game()
    demo = GameAssist(wdname)
    
    #level = int(input("请输入关数："))
    level = 1 

    while 1:
        print("level =", level)
        if level == 14:
            print("game over.")
            sys.exit()

        if not demo.start(game, level):
            print("game exit.")
            sys.exit()

        time.sleep(5)
        demo.mouse.click(820, 820)
        level = level + 1
        time.sleep(2)

        #level = int(input("请输入关数："))

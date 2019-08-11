

def open_file():

    """ 使用with语法打开一个文件 """
    try:
        f = open('../static/test.txt', 'r')
        rest = f.read()
        print(rest)
    except:
        pass
    finally:
        f.close()


    # 使用with不用手动关闭文件
    # with open('../static/test.txt', 'r') as f:
    #     rest = f.read()
    #     print(rest)




if __name__ == '__main__':
    open_file()
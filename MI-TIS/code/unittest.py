import unittest
import main

# 类名其实可以随意，只要继承自unittest.TestCase即可
# 但为了标识这个测试类是测试哪个类的，那就用Test+被测试类名
class TestBeTestClass(unittest.TestCase):

    # 只要是test_开头即可，但为了标识该测试方法是测试哪个方法的，就用test_+被测试方法[+ 数字]形式
    def test_my_add_1(self):
        print('execute test_one')
        obj = main.read.creat_excel_allfile_TIS_MI()
        assert obj == 1

if __name__ == '__main__':
    unittest.main()
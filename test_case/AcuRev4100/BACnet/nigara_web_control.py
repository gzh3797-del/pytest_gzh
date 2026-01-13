import logging
import time
import traceback
import pandas
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By


class NigaraWebControl():
    def __init__(self, url, username, password):
        # 添加option，使得网页打开后不关闭
        opt = Options()
        opt.add_experimental_option('detach', True)
        self.driver = webdriver.ChromiumEdge(opt)
        self.url = url
        self.username = username
        self.password = password
        self.actions = ActionChains(self.driver)

    def login(self):
        """
        函数简要描述:nigara 控制器登录函数
        详细描述:调用当前类时，传入url,用户名,密码，在当前函数进行登录
        参数：
        param1:当前类
        return:登录nigara控制器的页面，没有返回
        ExceptionType:登录失败,元素未找到或超时
        """
        self.driver.get(self.url)
        try:
            # 输入用户名
            user_input = self.driver.find_element(By.CLASS_NAME, 'login-input')
            user_input.send_keys(self.username)
            # 点击login button
            login_button = self.driver.find_element(By.ID, 'login-submit')
            login_button.click()
            time.sleep(2)
            # 输入密码
            pwd_input = self.driver.find_element(By.ID, 'password')
            pwd_input.send_keys(self.password)
            # 点击login button
            login_button = self.driver.find_element(By.ID, 'login-submit')
            login_button.click()
            time.sleep(10)
        except (NoSuchElementException, TimeoutException) as e:
            logging.error(f"登录失败:{e}")
            traceback.print_exc()
        except Exception as e:
            logging.error("登录失败，未知异常:", e)
            traceback.print_exc()

    def discover_device(self):
        """
        函数简要描述:nigara 控制器查找设备
        详细描述:已经登录的前提下，打开网页侧边栏，查找网络中的BACnet设备,当前函数只能查找1个设备
        参数：
        param1:当前类的Drive
        return:完成页面操纵，查找BACnet设备，没有返回
        ExceptionType:网页元素未找到或超时
        """
        try:
            # 点击navTreeSideBar,打开侧边栏
            navTreeSide_icon = self.driver.find_element(By.CSS_SELECTOR, '.inner.icon-icons-x16-listView')
            navTreeSide_icon.click()
            time.sleep(5)
            # 点击config旁边的三角符号，打开config菜单
            config_button = self.driver.find_element(By.XPATH,
                                                     '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/button')
            config_button.click()
            time.sleep(5)
            # config菜单打开后，点击Drivers旁边的三角符号，打开Drivers菜单
            drivers_button = self.driver.find_element(By.XPATH,
                                                      '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/ul/li[2]/button')
            drivers_button.click()
            time.sleep(5)
            # 双击 BacnetNetwork
            BacnetNetwork = self.driver.find_element(By.XPATH,
                                                     '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/ul/li[2]/ul/li[2]/span/label')
            self.actions.double_click(BacnetNetwork).perform()
            time.sleep(5)
            # 检查当前DataBase中，如果已有设备，进行删除
            database_object_num = self.driver.find_element(By.XPATH,
                                                           '//*[@id="bajaux"]/div[2]/div/div[2]/div/div[3]/div[1]/span[3]')
            object_num = database_object_num.text.split(' ')[0]
            print('object num:{}'.format(object_num))
            if int(object_num) != 0:
                for i in range(int(object_num)):
                    device_meau = self.driver.find_element(By.XPATH,
                                                           '//*[@id="bajaux"]/div[2]/div/div[2]/div/div[3]/div[2]/div/table/tbody/tr')
                    self.actions.context_click(device_meau).perform()
                    time.sleep(2)
                    delete = self.driver.find_element(By.XPATH, '/html/body/ul/li[14]')
                    delete.click()
                    time.sleep(2)
                    OK_button = self.driver.find_element(By.XPATH, '/html/body/div[2]/div/ul/li[1]/button')
                    OK_button.click()
                    time.sleep(2)
            # 点击Discover 按钮，查找设备
            discover_button = self.driver.find_element(By.XPATH, '//*[@id="bajaux"]/div[2]/div/div[3]/button[4]')
            discover_button.click()
            time.sleep(1)
            OK_button = self.driver.find_element(By.XPATH, '/html/body/div[2]/div/ul/li[1]/button')
            OK_button.click()
            time.sleep(20)
        except (NoSuchElementException, TimeoutException) as e:
            logging.error(f"登录失败:{e}")
            traceback.print_exc()
        except Exception as e:
            logging.error("登录失败，未知异常:", e)
            traceback.print_exc()

    def add_device(self):
        """
        函数简要描述:nigara 控制器添加设备
        详细描述:已经查找完毕设备的前提下，添加第一个BACnet设备
        参数：
        param1:当前类的Drive
        return:完成页面操纵，添加BACnet设备，没有返回
        ExceptionType:网页元素未找到或超时
        """
        try:
            add_device_button = self.driver.find_element(By.XPATH,
                                                         '//*[@id="bajaux"]/div[2]/div/div[2]/div/div[1]/div/div[2]/div/table/tbody/tr')
            self.actions.double_click(add_device_button).perform()
            time.sleep(1)
            OK_button = self.driver.find_element(By.XPATH, '/html/body/div[2]/div/ul/li[1]/button')
            OK_button.click()
            time.sleep(5)
        except (NoSuchElementException, TimeoutException) as e:
            logging.error(f"登录失败:{e}")
            traceback.print_exc()
        except Exception as e:
            logging.error("登录失败，未知异常:", e)
            traceback.print_exc()

    def discover_points(self):
        """
        函数简要描述:nigara 控制器查找已添加设备的points
        详细描述:已经查找完毕设备的前提下，添加第一个BACnet设备
        参数：
        param1:当前类的Drive
        return:完成页面操纵，查找设备的points，没有返回
        ExceptionType:网页元素未找到或超时
        """
        try:
            # 点击BacnetNetwork菜单左侧的三角按钮，打开菜单
            BacnetNetwork_button = self.driver.find_element(By.XPATH,
                                                            '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/ul/li[2]/ul/li[2]/button')
            BacnetNetwork_button.click()
            time.sleep(5)
            # 点击4110 设备菜单左侧的三角按钮，打开菜单
            device_button_4110 = self.driver.find_element(By.XPATH,
                                                          '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/ul/li[2]/ul/li[2]/ul/li[5]/button')
            device_button_4110.click()
            time.sleep(5)
            # 双击Points 菜单
            points = self.driver.find_element(By.XPATH,
                                              '//*[@id="WebShell-navTreeSideBar"]/div[2]/ul/li[1]/ul/li[2]/ul/li[2]/ul/li[5]/ul/li[2]/span/label')
            self.actions.double_click(points).perform()
            time.sleep(5)
            # Points页面点击Discover，查找参数
            discover_button = self.driver.find_element(By.XPATH, '//*[@id="bajaux"]/div[2]/div/div[2]/button[4]')
            discover_button.click()
            # 这里的死等可以优化成获取页面的元素，success之后退出
            for i in range(300):
                time.sleep(1)
                success_info = self.driver.find_element(By.XPATH,
                                                        '//*[@id="bajaux"]/div[2]/div/div[1]/div/div[1]/div/div[1]/div/div[2]/div[1]')
                if success_info.text == 'Success':
                    time.sleep(10)
                    break
        except (NoSuchElementException, TimeoutException) as e:
            logging.error(f"登录失败:{e}")
            traceback.print_exc()
        except Exception as e:
            logging.error("登录失败，未知异常:", e)
            traceback.print_exc()

    def points_info_output(self, output_path):
        """
        函数简要描述:nigara 控制器解析已经完成查找的全部points，并输出到excel文件
        详细描述:已经查找完毕points的前提下，获取页面上的全部设备points，解析输出到excel文件
        参数：
        param1:当前类的Drive，输出结果文件的路径
        return:解析网页中查找到的全部points，输出到excel文件
        ExceptionType:网页元素未找到或超时
        """
        # 获取所有points，并解析
        try:
            object_table = self.driver.find_element(By.CLASS_NAME, 'ux-table')
            rows = object_table.find_elements(By.XPATH,
                                              '//*[@id="bajaux"]/div[2]/div/div[1]/div/div[1]/div/div[3]/div/table/tbody/tr')

            object_table_data = []
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                row_data = [cell.text.strip() for cell in cells]
                object_table_data.append(row_data)

            logging.info("表格内容:")
            for row in object_table_data:
                logging.info(row)
            columns = ['Object Name', 'Object Name', 'Object ID', 'Property ID', 'Index', 'Value', 'Description']
            object_table_data_xl = pandas.DataFrame(object_table_data, columns=columns)
            object_table_data_xl.to_excel(output_path, index=False)
        except NoSuchElementException as e:
            logging.error(f"元素查找失败: {e}")
        except TimeoutException as e:
            logging.error(f"网页加载超时: {e}")
        except FileNotFoundError as e:
            logging.error(f"文件路径错误: {e}")
        except PermissionError as e:
            logging.error(f"文件权限错误: {e}")
        except Exception as e:
            logging.error(f"其他错误: {e}")

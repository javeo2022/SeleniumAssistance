# -*- coding: utf-8 -*-
import logging
import os
import shutil
import subprocess
import tempfile
from subprocess import CREATE_NO_WINDOW
from time import sleep

from bs4 import BeautifulSoup
from selenium import webdriver
# from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common import utils
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
# from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
# from selenium.webdriver.support.ui import WebDriverWait
# from webdriver_manager.chrome import ChromeDriverManager


def get_module_logger():
    logger = logging.getLogger(os.path.splitext(os.path.basename(__file__))[0])
    logger = _set_handler(logger, logging.StreamHandler())
    logger.setLevel(logging.DEBUG)
    # logger.setLevel(logging.INFO)
    # logger.setLevel(logging.ERROR)
    logger.propagate = False
    return logger


def _set_handler(logger, handler):
    handler.setFormatter(logging.Formatter('%(asctime)s %(name)s:%(lineno)s %(funcName)s [%(levelname)s]: %(message)s'))
    logger.addHandler(handler)
    return logger


logger = get_module_logger()


class ScrapingAssistance():
    def __init__(self) -> None:
        self.selenium_flg = False
        self.bs_flg = False

    def __del__(self):
        if self.selenium_flg is True:
            self.selenium_quit()

    def selenium_open(self, headless: bool = False, user_data_dir: str = '', timeout: int = 10, imagesEnabled: bool = True,
                      remote_debugging_mode: bool = False, port: int = 0, driver_path: str = '', mobile_mode: str = '',
                      download_dir: str = ''):
        '''
        headless:ヘッドレスモードで実行するか
        user_data_dir:ユーザープロファイルを指定する場合フォルダパスを指定※"Default"か"Profile x"を指定
        timeout:デフォルトのタイムアウト値
        remote_debugging_mode:chromeをchromedriverで操作する？ ※CloudFlare対策になる
        port:remote_debugging_modeを有効にした時のポート番号 ※指定不要
        driver_path:指定のchromedriver.exeがある時にフルパスを指定
        mobile_mode:"android" or "ios"を指定するとレスポンシブで開く
        '''
        self.timeout = timeout
        options = Options()
        # 引数のdriver_pathに指定がないor指定のパスに誤りがあれば最新化しつつパスを決める
        # 判定のためにパス情報を正規化しておく　※''は'.'になるので注意
        driver_path = os.path.normpath(driver_path)
        if driver_path == '.':
            # 指定がない場合
            driver_path = ''
        elif os.path.isfile(driver_path) is False:
            # 指定があるけどファイルがない場合
            driver_path = ''
        elif os.path.basename(driver_path) != 'chromedriver.exe':
            # ファイルはあったけど'chromedriver.exe'じゃない場合
            driver_path = ''

        # 引数のuser_data_dirに指定がないならプロファイル情報をtmpに作る※最後にフォルダ削除する判定も付ける
        # 判定のためにパス情報を正規化しておく　※''は'.'になるので注意
        self.user_data_dir = os.path.normpath(user_data_dir)
        # 指定がないorプロファイルのフォルダ名（DefaultかProfileで始まる）ではない場合はテンポラリにフォルダ作成して指定する
        if self.user_data_dir == '.' or (self.user_data_dir != 'Default' and self.user_data_dir.startswith('Profile') is False):
            self.user_data_dir = os.path.normpath(tempfile.mkdtemp())
            self.delete_user_data_dir = self.user_data_dir
            self.user_data_dir = os.path.join(self.user_data_dir, 'User Data', 'Default')
            os.makedirs(self.user_data_dir, exist_ok=True)
            # 最後にプロファイルデータを削除する
            self.user_data_delete = True
        else:
            # 最後にプロファイルデータを削除しない
            self.user_data_delete = False
        options.add_argument('--user-data-dir={}'.format(self.user_data_dir))

        # 起動オプションの引数を設定する　※お好みでコメントアウトを
        options.add_argument('--disable-gpu')                   # GPUハードウェアアクセラレーションを無効
        options.add_argument('--disable-extensions')            # 全ての拡張機能を無効
        options.add_argument('--proxy-server="direct://"')      # Proxy経由ではなく直接接続
        options.add_argument('--proxy-bypass-list=*')           # すべてのホスト名
        options.add_argument('--start-maximized')               # 初期のウィンドウ最大化
        # options.add_argument('--no-sandbox')                    # Chromeの保護機能を無効
        options.add_argument('--incognito')                     # Chromeをシークレットモードで起動
        prefs = {
            'profile.default_content_setting_values.notifications': 2,                 # 通知ポップアップを無効
            # 'profile.managed_default_content_settings.javascript': 2,                  # Javascriptを無効
            'credentials_enable_service': False,                                       # パスワード保存のポップアップを無効
            'profile.password_manager_enabled': False,                                 # パスワード保存のポップアップを無効
            # 'download.default_directory' : download_dir                               # ダウンロード先のディレクトリを指定
        }
        options.add_experimental_option('prefs', prefs)

        # ヘッドレスモードの場合
        if headless is True:
            options.add_argument('--headless=new')                  # ヘッドレスモードで起動する
        if imagesEnabled is False:
            options.add_argument('--blink-settings=imagesEnabled=false')    # 画像を読み込まない

        # スマホ表示にする場合のオプション
        if mobile_mode == 'android':
            mobile_emulation = {'deviceName': 'Galaxy S9+'}
            options.add_experimental_option('mobileEmulation', mobile_emulation)
        elif mobile_mode == 'ios':
            # 最近iOS番にならない
            mobile_emulation = {'deviceName': 'iPhone SE'}
            options.add_experimental_option('mobileEmulation', mobile_emulation)

        # ダウンロード先フォルダが指定されていれば
        if download_dir != '':
            if os.path.isdir(download_dir) is True:
                prefs = {'download.default_directory': download_dir}
                options.add_experimental_option('prefs', prefs)

        # chrome.exeをデバックで操作する
        if remote_debugging_mode is True:
            # ポート番号を設定する
            debug_port = str(port if port != 0 else utils.free_port())
            # ローカルホスト
            debug_host = '127.0.0.1'
            options.add_argument('--remote-debugging-host={}'.format(debug_host))
            options.add_argument('--remote-debugging-port={}'.format(debug_port))
            # options.debugger_address = '{}:{}'.format(debug_host, debug_port)
            # chrome.exeの在りかを探す
            chrome_exe_path = find_chrome_executable()
            self.driver = subprocess.Popen([chrome_exe_path, *options.arguments], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.browser_pid = self.driver
        else:
            self.service = Service()
            self.service.creation_flags = CREATE_NO_WINDOW
            self.driver = webdriver.Chrome(service=self.service, options=options)

        # ページロードの最大待機時間
        # self.driver.set_page_load_timeout(timeout)
        # 要素が見つかるまでの最大待機時間
        # self.driver.implicitly_wait(timeout)
        # js完了までの最大待機時間
        # self.driver.set_script_timeout(timeout)
        # とりあえずGoogleを開く
        self.driver.get('https://www.google.com/?hl=ja')
        self.selenium_flg = True

    def selenium_quit(self):
        # 閉じる時はkillでプロセスを落とす
        self.driver.quit()
        try:
            self.service.process.kill()
        except Exception:
            pass
        try:
            os.kill(self.browser_pid, 15)
        except Exception:
            pass
        # 一時的に作ったプロファイルデータのフォルダを削除する
        try:
            if self.user_data_delete is True:
                shutil.rmtree(self.delete_user_data_dir)
        except Exception:
            pass

    def selenium_input(self, by: By, value: str, str: str = '', idx: int = 0, timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        str:実際に入力する文字
        idx:インデックスの指定が必要な場合 ※初期値は0
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return
            0:正常終了
            1:指定した要素がDOM上にない
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            99:その他
        '''
        if timeout_second == -1:
            timeout_second = self.timeout

        # 指定したインデックスが要素の数より多い場合DOM上にない判定
        if len(self.driver.find_elements(by, value)) <= idx:
            return {'rc': 1, 'message': 'This element is not found'}

        return self.selenium_input_elm(self.driver.find_elements(by, value)[idx], str, timeout_second)

    def selenium_input_elm(self, element: WebElement, str: str = '', timeout_second: int = -1) -> dict:
        '''
        WebElementをダイレクトに指定して文字入力をする
        element:指定するWebElementを設定
        str:実際に入力する文字
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return:
            0:正常終了
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            99:その他
        '''
        # タイムアウト秒数を設定する
        if timeout_second == -1:
            timeout_second = self.timeout

        # 画面上に表示されているか確認する
        if element.is_displayed is False:
            return {'rc': 2, 'message': 'This element is not visible'}

        # タグがinputでもtextareaでもなければエラー扱いにする
        if element.tag_name != 'input' and element.tag_name != 'textarea':
            return {'rc': 3, 'message': 'This element\'s tagName is not input and textarea'}

        # タグが入力可能か確認する
        if element.is_enabled() is False:
            return {'rc': 4, 'message': 'This element is denabled'}

        # 念のため画面をelementへ動かしておく
        ActionChains(self.driver).move_to_element(element).perform()

        # ここから正常時の操作
        element.clear       # send_keysだと追記になってしまうので一旦クリアしておく
        element.click()     # 念のためクリックしてアクティブにする
        sleep(1)
        element.send_keys(str)          # ここで実際に文字入力する
        element.send_keys(Keys.TAB)     # 入力後に発火するjsの可能性があるのでタブ移動する

        sleep(1)
        return {'rc': 0, 'message': 'OK'}

    def selenium_get(self, by: By, value: str, idx: int = 0, att: str = 'default', timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        int:インデックスの指定が必要な場合 ※初期値は0
        att:取得する対象 ※textかattributeに指定する値
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return ['rc']:
            0:正常終了
            1:指定した要素がDOM上にない
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            99:その他
        return ['text']:
            異常終了した場合も''を返す
        '''
        # 指定したインデックスが要素の数より多い場合DOM上にない判定
        if len(self.driver.find_elements(by, value)) <= idx:
            return {'rc': 1, 'text': '', 'message': 'This element is not found'}

        return self.selenium_get_elm(self.driver.find_elements(by, value)[idx], att, timeout_second)

    def selenium_get_elm(self, element: WebElement, att: str = 'default', timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        str:実際に入力する文字
        int:インデックスの指定が必要な場合 ※初期値は0
        return ['rc']:
            0:正常終了
            2:画面上に見えない
            4:入力可能状態じゃない
            5:指定したattributeがない
            99:その他
        return ['text']:
            異常終了した場合も''を返す
        '''
        # タイムアウト秒数を設定する
        if timeout_second == -1:
            timeout_second = self.timeout

        # 画面上にelementが表示されているか確認する
        if element.is_displayed is False:
            return {'rc': 2, 'text': '', 'message': 'This element is not visible'}

        # タグが入力可能か確認する
        if element.is_enabled() is False:
            return {'rc': 4, 'text': '', 'message': 'This element is denabled'}

        # 念のため画面をelementへ動かしておく
        ActionChains(self.driver).move_to_element(element).perform()

        # 取得対象に指定がない場合ははtextを確認して値がない場合はvalueを確認する
        if att == 'default':
            ret = element.text
            if len(ret.strip()) == 0:
                ret = element.get_attribute('value')
        elif att == 'text':
            # 明示的にtextが指定された場合はtextしか確認しない
            ret = element.text
        else:
            # text以外はattの指定にする
            ret = element.get_attribute(att)

        # 返値がNoneはattrに該当がなかった場合
        if ret is None:
            return {'rc': 5, 'text': '', 'message': 'This element\'s attribute is not found'}

        sleep(1)
        return {'rc': 0, 'message': 'OK'}

    def selenium_click(self, by: By, value: str, idx: int = 0, timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        str:実際に入力する文字
        idx:インデックスの指定が必要な場合 ※初期値は0
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return
            0:正常終了
            1:指定した要素がDOM上にない
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            99:その他
        '''
        # 指定したインデックスが要素の数より多い場合DOM上にない判定
        if len(self.driver.find_elements(by, value)) <= idx:
            return {'rc': 1, 'message': 'This element is not found'}

        return self.selenium_click_elm(self.driver.find_elements(by, value)[idx], timeout_second)

    def selenium_click_elm(self, element: WebElement, timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        int:インデックスの指定が必要な場合 ※初期値は0
        return
            0:正常終了
            1:DOM上に要素が見つからない
            2:画面上に見えない
            3:クリックできる状態じゃない
            4:入力可能状態じゃない
            99:その他
        '''
        # タイムアウト秒数を設定する
        if timeout_second == -1:
            timeout_second = self.timeout

        # 画面上にelementが表示されているか確認する
        if element.is_displayed is False:
            return {'rc': 2, 'message': 'This element is not visible'}

        # タグが入力可能か確認する
        if element.is_enabled() is False:
            return {'rc': 4, 'message': 'This element is denabled'}

        # 念のため画面をelementへ動かしておく
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        ActionChains(self.driver).move_to_element(element).perform()
        element.click()

        sleep(1)
        return {'rc': 0, 'message': 'OK'}

    def selenium_select(self, by: By, value: str, idx: int = 0, target_type: str = 'visible_text', target_value: str = '', timeout_second: int = -1) -> dict:
        '''
        by:指定するByを設定
        value:byで指定する文字列
        idx:インデックスの指定が必要な場合 ※初期値は0
        target_type:"index"、"value"、"visible_text"のいずれか
        target_value:target_typeに合わせて選択する値
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return
            0:正常終了
            1:指定した要素がDOM上にない
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            6:タグがselectじゃない
            99:その他
        '''
        # 指定したインデックスが要素の数より多い場合DOM上にない判定
        if len(self.driver.find_elements(by, value)) <= idx:
            return {'rc': 1, 'message': 'This element is not found'}

        return self.selenium_select_elm(self.driver.find_elements(by, value)[idx], target_type, target_value, timeout_second)

    def selenium_select_elm(self, element: WebElement, target_type: str = 'visible_text', target_value: str = '', timeout_second: int = -1) -> dict:
        '''
        WebElementをダイレクトに指定して文字入力をする
        element:指定するWebElementを設定
        target_type:"index"、"value"、"visible_text"のいずれか
        target_value:target_typeに合わせて選択する値
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return:
            0:正常終了
            2:画面上に見えない
            3:タグがinputかtextarea以外
            4:入力可能状態じゃない
            6:タグがselectじゃない
            7:target_typeの入力ミス
            8:target_valueに該当がない
            99:その他
        '''
        # タイムアウト秒数を設定する
        if timeout_second == -1:
            timeout_second = self.timeout

        # 画面上にelementが表示されているか確認する
        if element.is_displayed is False:
            return {'rc': 2, 'message': 'This element is not visible'}

        # タグが変更可能か確認する
        if element.is_enabled() is False:
            return {'rc': 4, 'message': 'This element is denabled'}

        # タグを確認する
        if element.tag_name.lower() != 'select':
            return {'rc': 6, 'message': 'This element\'s tagName is not select'}

        # 念のため画面をelementへ動かしておく
        ActionChains(self.driver).move_to_element(element).perform()

        # ここから正常時の操作
        select = Select(element)
        for idx, select_option in enumerate(select.options):
            if target_type == 'index':
                if str(idx) == target_value:
                    select.select_by_index(target_value)
                    break
            elif target_type == 'value':
                if select_option.get_attribute("value") == target_value:
                    select.select_by_value(target_value)
                    break
            elif target_type == 'visible_text':
                if select_option.text == target_value:
                    select.select_by_visible_text(target_value)
                    break
            else:
                return {'rc': 7, 'message': 'target_type is input error'}
        else:
            return {'rc': 8, 'message': 'target_value is not applicable'}

        sleep(1)
        return {'rc': 0, 'message': 'OK'}

    def bs4_set(self, page_source: str = '', parser: str = 'html.parser'):
        '''
        int:インデックスの指定が必要な場合 ※初期値は0
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return[0]:
            0:正常終了
            1:指定した要素がDOM上にない
            5:指定したattributeがない
            99:その他
        return[1]:
            異常終了した場合も''を返す
        '''
        if page_source == '':
            page_source = self.driver.page_source

        self.soup = BeautifulSoup(page_source, parser)
        self.bs_flg = True

    def bs4_get(self, selector: str = '', target_attribute: str = 'text', idx: int = 0) -> dict:
        '''
        int:インデックスの指定が必要な場合 ※初期値は0
        timeout_second:タイムアウトまでの秒数 ※初期値は別途設定
        return[0]:
            0:正常終了
            1:指定した要素がDOM上にない
            5:指定したattributeがない
            99:その他
        return[1]:
            異常終了した場合も''を返す
        '''
        if self.bs_flg is False:
            return {'rc': 9, 'text': '', 'message': 'Not set BeautifulSoup'}

        # 指定したインデックスが要素の数より多い場合DOM上にない判定
        if len(self.soup.select(selector)) <= idx:
            return {'rc': 1, 'text': '', 'message': 'This selector is not found'}

        if target_attribute == 'text':
            ret = self.soup.select(selector)[idx].get_text()
        else:
            ret = self.soup.select(selector)[idx].get(target_attribute)

        if ret is None:
            return {'rc': 5, 'text': '', 'message': 'This selector\'s attribute is not found'}

        return {'rc': 0, 'text': ret, 'message': 'OK'}

    def bs4_return_select(self, selector: str = ''):
        return self.soup.select(selector=selector)

    def window_scroll(self):
        '''
        js対応のためにゆっくり画面を一番下にスクロールする
        '''
        win_height = self.driver.execute_script('return window.innerHeight')
        Top = 1
        last_height = self.driver.execute_script('return document.body.scrollHeight')
        while Top < last_height:
            Top = Top + int(win_height * 0.8)
            self.driver.execute_script(f'window.scrollTo(0, "{str(Top)}")')
            last_height = self.driver.execute_script('return document.body.scrollHeight')
            sleep(1)


def find_chrome_executable():
    '''
    chrome.exe のフルパスを取得する
    '''
    candidates = set()
    for item in map(os.environ.get, ('PROGRAMFILES', 'PROGRAMFILES(X86)', 'LOCALAPPDATA', 'PROGRAMW6432'),):
        if item is not None:
            for subitem in ('Google/Chrome/Application', 'Google/Chrome Beta/Application', 'Google/Chrome Canary/Application',):
                candidates.add(os.sep.join((item, subitem, 'chrome.exe')))

    for candidate in candidates:
        if os.path.exists(candidate) and os.access(candidate, os.X_OK):
            return os.path.normpath(candidate)

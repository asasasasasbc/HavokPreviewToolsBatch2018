print('---Core hkx batch convert program 1.2 by Forsakensilver---')
print('---hkx批量转换工具 1.2 by 遗忘的银灵---')
#一个奇怪的bug，见https://github.com/pywinauto/pywinauto/issues/517
import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED

import os, time
import glob
from tkinter import Tk, filedialog
from pywinauto import Desktop, Application
from pywinauto.findwindows import find_elements
from pywinauto.keyboard import send_keys
from pywinauto.timings import wait_until

import configparser

# 创建ConfigParser对象
config = configparser.ConfigParser()

# 读取INI文件
config.read('./config.ini',encoding='utf-8')

# 获取所有的section
sections = config.sections()
#print('Sections:', sections)

# 读取特定section的配置项
FORMAT_OPTION = 'MSVC x64, XBoxOne (64m)'
FORMAT_OPTION = config.get('settings', 'format_option').strip("'")

IN_FILE_SUFFIX = config.get('settings', 'in_file_suffix').strip("'")
OUT_FILE_SUFFIX = config.get('settings', 'out_file_suffix').strip("'")
ENTER_KEY = config.get('settings', 'enter_key').strip("'")

# 打印获取的配置项
print('FORMAT_OPTION:', FORMAT_OPTION)

def convert_file(src_file_path, out_file_path):
    #找到PREVIEWTOOL地址
    #APP_PATH = "D:\\Program Files\\Havok\\HavokContentTools\\PreviewTool.exe"
    #SOURCE_FILE_PATH = "G:\\DarkSoulsMod\\AnimationTweak\\a000_010000.hkx"
    #OUT_FILE_PATH = 'G:\\DarkSoulsMod\\AnimationTweak\\a000_010000_out.hkt'
    #FORMAT_OPTION = 'MSVC x64, XBoxOne (64m)'

    SOURCE_FILE_PATH = src_file_path
    OUT_FILE_PATH = out_file_path
    #FORMAT_OPTION = 'MSVC x64, XBoxOne (64m)'
    print(f"Converting {src_file_path} to {out_file_path}")
    #app = Application().connect(path=APP_PATH) #早期的方法
    app = Application().connect(title="Havok Preview Tool [StandAlone]")
    if not app:
        print('无法找到正在运行的havok程序！请先启动previewTools.exe!\nPlease start previewTools.exe first!')
    # 如果找到了，连接到第一个实例
    if app:
        # 获取主窗口
        main_window = app.window(title="Havok Preview Tool [StandAlone]")

        # 对窗口进行操作
        #main_window.Edit.type_keys(" ", with_newlines=True)
        # 激活窗口
        main_window.set_focus()

        # 等待直到窗口真正被激活，最多等待5秒
        wait_until(5,1, lambda: main_window.is_active())
        #------------------------------------------------
        # 选择并点击 "文件" 菜单中的 "打开" 选项
        #main_window.menu_select("File->Open...") 
        # 发送快捷键打开 "文件" 菜单
        send_keys('%F')  # Alt+F 打开 "文件" 菜单
        send_keys('O')   # O 打开 "打开" 对话框
        #-------------------------------------------------

        #-------------------------------------------------
        # 使用窗口的标题来获取文件选择对话框
        print('Finding opening dialog...')
        dialog = app.window(title='Select a Havok packfile/tagfile to load into source scene')

        # 等待对话框完全加载
        dialog.wait('ready', timeout=5)
        dialog.set_focus()
        # 输入文件路径，这里需要您提供完整的文件路径
        file_path = SOURCE_FILE_PATH
        edit = dialog.child_window(class_name='Edit')
        edit.set_text(file_path)
        #send_keys('%F')  # Alt+F 打开 "文件" 菜单

        send_keys('%O')  # Alt+o 触发文件打开按钮
        #-------------------------------------------------

        #有些时候导入可能会报错，这边加按个enter按钮
        if ENTER_KEY == 'yes':
            time.sleep(0.5)
            send_keys("{ENTER}")
            time.sleep(0.5)

        send_keys('%F')  # Alt+F 打开 "文件" 菜单
        send_keys('S')   # S 打开 "保存" 对话框
        #-------------------------------------------------
        #保存文件
        dialog = app.window(title='Save To File')
        # 定位到ComboBox控件，这里需要根据实际情况替换控件的识别属性
        combo_box = dialog.child_window(auto_id='formatComboBox')
        #这儿可以改成其他的选项，比方说
        combo_box.select(FORMAT_OPTION)
        #或者XML (Cross Platform Text)
        edit = dialog.child_window(auto_id='filenameTextBox')
        out_file_path = OUT_FILE_PATH
        #if out_file_path[-1] == 'x': #hkx >hkt
        #    out_file_path = out_file_path[0:len(out_file_path)-1] + 't'
        #    print('Save to hkt:' + out_file_path)
        # 检查原字符串是否以旧的结尾结尾
        if out_file_path.endswith(IN_FILE_SUFFIX):
            # 使用切片去掉旧的结尾，然后加上新的结尾
            out_file_path = out_file_path[:-len(IN_FILE_SUFFIX)] + OUT_FILE_SUFFIX
            edit.set_text(out_file_path)

        okButton = dialog.child_window(auto_id='okButton')
        okButton.click()
        print('Done.')

def main():
    
    #time.sleep(1)
    print('请选择包含hkx的文件夹路径,并确保在此之前先打开运行previewTool.exe！')
    print('Please select the folder path containing "hkx" and make sure to run "previewTool.exe" before proceeding!')
    folder_path = filedialog.askdirectory(title='Please choose the folder that contains havok files.')
    if not folder_path:
        print("没有选择文件夹。\nHaven't chosen folder yet.")
        return
    
    print('下一步...')

    hkx_files = glob.glob(os.path.join(folder_path, '*.' + IN_FILE_SUFFIX))
    if not hkx_files:
        print(f"在选定的文件夹中没有找到.{IN_FILE_SUFFIX}文件。\nCannot find .{IN_FILE_SUFFIX}file in the chosen folder.")
        return

    output_folder_path = os.path.join(folder_path, 'Output')
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    for hkx_file in hkx_files:
        filename = os.path.basename(hkx_file)
        out_file = os.path.join(output_folder_path, filename).replace('/','\\')
        in_file = hkx_file.replace('/','\\')
        convert_file(in_file, out_file)
    print('All jobs done.')

if __name__ == "__main__":
    main()

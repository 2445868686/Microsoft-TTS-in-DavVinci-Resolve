import platform
from xml.dom import minidom
import xml.etree.ElementTree as ET
import sys
import os
import time
import webbrowser
import json
import azure.cognitiveservices.speech as speechsdk
try:
    import DaVinciResolveScript as dvr_script
    from python_get_resolve import GetResolve
    print("DaVinciResolveScript from Python")
except ImportError:
    
    if platform.system() == "Darwin": 
        resolve_script_path1 = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Examples"
        resolve_script_path2 = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules"
    elif platform.system() == "Windows": 
        resolve_script_path1 = os.path.join(os.environ['PROGRAMDATA'], "Blackmagic Design", "DaVinci Resolve", "Support", "Developer", "Scripting", "Examples")
        resolve_script_path2 = os.path.join(os.environ['PROGRAMDATA'], "Blackmagic Design", "DaVinci Resolve", "Support", "Developer", "Scripting", "Modules")
    else:
        raise EnvironmentError("Unsupported operating system")

    sys.path.append(resolve_script_path1)
    sys.path.append(resolve_script_path2)

    try:
        import DaVinciResolveScript as dvr_script
        from python_get_resolve import GetResolve
        print("DaVinciResolveScript from DaVinci")
    except ImportError as e:
        raise ImportError("Unable to import DaVinciResolveScript or python_get_resolve after adding paths") from e

def check_or_create_file(file_path):
    if os.path.exists(file_path):
        pass
    else:
        try:
            with open(file_path, 'w') as file:
                json.dump({}, file)  # 写入一个空的JSON对象，以初始化文件
        except IOError:
            raise Exception(f"Cannot create file: {file_path}")

def get_script_path():
    # 使用 sys.argv 获取脚本路径
    if len(sys.argv) > 0:
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        return os.getcwd()

script_path = get_script_path()

settings_file = os.path.join(script_path, 'TTS_settings.json')

check_or_create_file(settings_file)

def load_settings(settings_file):
    if os.path.exists(settings_file):
        with open(settings_file, 'r') as file:
            content = file.read()
            if content:
                try:
                    settings = json.loads(content)
                    return settings
                except json.JSONDecodeError as err:
                    print('Error decoding settings:', err)
                    return None
    return None

def save_settings(settings, settings_file):
    with open(settings_file, 'w') as file:
        content = json.dumps(settings, indent=4)
        file.write(content)

saved_settings = load_settings(settings_file) 

default_settings = {
    "API_KEY": '',
    "OUTPUT_DIRECTORY": '',
    "REGION": '',
    "LANGUAGE": 0,
    "TYPE": 0,
    "NAME": 0,
    "RATE": 1.0,
    "PITCH": 1.0,
    "VOLUME": 1.0,
    "STYLE": 0,
    "BREAKTIME":50,
    "STYLEDEGREE": 1.0,
    "OUTPUT_FORMATS":0
}


# 获取 DaVinci Resolve 脚本接口
resolve = GetResolve()
projectManager = resolve.GetProjectManager()
currentProject = projectManager.GetCurrentProject()
currentTimeline = currentProject.GetCurrentTimeline()

def add_to_media_pool(filename):
    # 获取Resolve实例
    resolve = dvr_script.scriptapp("Resolve")
    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()
    media_pool = project.GetMediaPool()
    root_folder = media_pool.GetRootFolder()
    tts_folder = None

    folders = root_folder.GetSubFolderList()
    for folder in folders:
        if folder.GetName() == "TTS":
            tts_folder = folder
            break

    if not tts_folder:
        tts_folder = media_pool.AddSubFolder(root_folder, "TTS")

    if tts_folder:
        print(f"TTS folder is available: {tts_folder.GetName()}")
    else:
        print("Failed to create or find TTS folder.")
        return False

    media_storage = resolve.GetMediaStorage()
    mapped_path = filename
    media_pool.SetCurrentFolder(tts_folder)
    return media_pool.ImportMedia([mapped_path], tts_folder)
infomsg = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {
            font-family: Arial, sans-serif;
            color: #ffffff;
            padding: 20px;
        }
        h3 {
            font-weight: bold;
            font-size: 1.5em;
            margin-top: 15px;
            margin-bottom: 0px; /* 调整此处以减少间隔 */
            border-bottom: 2px solid #f0f0f0;
            padding-bottom: 5px;
            color: #c7a364; /* 黄色 */
        }
        p {
            font-size: 1.2em;
            margin-top: 5px;
            margin-bottom: 0px; /* 调整此处以减少间隔 */
            color: #a3a3a3; /* 白色 */
        }
        a {
            color: #1e90ff;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <h3>介绍</h3>
    <p>此脚本使用Azure的TTS功能将文字转换为语音。</p>

    <h3>保存路径</h3>
    <p>指定生成文件的保存路径。</p>

    <h3>区域、API密钥</h3>
    <p>从<a href="https://speech.microsoft.com/">Microsoft Speech Studio</a>获取您的API密钥。</p>
</body>
</html>

"""

"""

<body>
    <h3>Introduction</h3>
    <p>This script uses Azure's TTS feature to convert text to speech.</p>

    <h3>Save Path</h3>
    <p>Specifies the save path for the generated files.</p>

    <h3>Region、API Key</h3>
    <p>Obtain your API key from <a href="https://speech.microsoft.com/">Microsoft Speech Studio</a></p>
</body>
"""

ui = fusion.UIManager
dispatcher = bmd.UIDispatcher(ui)
win = dispatcher.AddWindow(
    {
        "ID": 'MyWin',
        "WindowTitle": 'TTS 1.0 赞助版',
        "Geometry": [700, 300, 425, 600],
        "Spacing": 10,
    },
 [
        ui.VGroup(
            [
                ui.TabBar({"Weight": 0.0, "ID": "MyTabs"}),
                ui.VGroup([]),
                ui.Stack(
                    {"Weight": 1.0, "ID": "MyStack"},
                    [
                        ui.VGroup(
                            {"Weight": 1},
                            [

                                ui.HGroup(
                                    {"Weight": 1},
                                    [
                                        ui.TextEdit({"ID": 'SubtitleTxt', "Text": '',"Weight": 1}),
                                    ]
                                ),

                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                       
                                        ui.SpinBox({"ID": 'BreakSpinBox', "Value": 50, "Minimum": 0, "Maximum": 5000, "SingleStep": 50, "Weight": 0.1}),
                                        ui.Label({"ID": 'BreakLabel', "Text": 'ms',  "Weight": 0.1}),
                                        ui.Button({"ID": 'BreakButton', "Text": 'Break', "Weight": 0.1}),
                                        ui.Button({"ID": 'GetSubButton', "Text": 'Get Subtitle From Timeline', "Weight": 1}),

                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        ui.Label({"ID": 'LanguageLabel', "Text": 'Language', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'LanguageCombo', "Text": '', "Weight": 0.8}),
                                        ui.Label({"ID": 'NameTypeLabel', "Text": 'Type', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'NameTypeCombo', "Text": '', "Weight": 0.8}),
                                    
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        ui.Label({"ID": 'NameLabel', "Text": 'Name', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'NameCombo', "Text": '', "Weight": 0.8}),
                                        ui.Label({"ID": 'MultilingualLabel', "Text": 'Multilingual', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'MultilingualCombo', "Text": '', "Weight": 0.2}),      
                                  
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'StyleLabel', "Text": 'Style', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'StyleCombo', "Text": '', "Weight": 0.8}),
                                        
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'StyleDegreeLabel', "Text": 'Style Degree', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'StyleDegreeSlider', "Value": 100, "Minimum": 0, "Maximum": 300,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'StyleDegreeSpinBox', "Value": 1.0, "Minimum": 0.0, "Maximum": 3.0, "SingleStep": 0.01, "Weight": 0.3}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'RateLabel', "Text": 'Rate', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'RateSlider', "Value": 100, "Minimum": 0, "Maximum": 300,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'RateSpinBox', "Value": 1.0, "Minimum": 0.0, "Maximum": 3.0, "SingleStep": 0.01, "Weight": 0.3}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'PitchLabel', "Text": 'Pitch', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'PitchSlider', "Value": 100, "Minimum": 50, "Maximum": 150,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'PitchSpinBox', "Value": 1.0, "Minimum": 0.5, "Maximum": 1.5, "SingleStep": 0.01, "Weight": 0.3}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'VolumeLabel', "Text": 'Volume', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'VolumeSlider', "Value": 100, "Minimum": 0, "Maximum": 150,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'VolumeSpinBox', "Value": 1.0, "Minimum": 0, "Maximum": 1.5, "SingleStep": 0.01, "Weight": 0.3}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        ui.Label({"ID": 'OutputFormatLabel', "Text": 'Format', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'OutputFormatCombo', "Text": 'Output_Format', "Weight": 0.8}),
                                    ]
                                ),

                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        ui.Button({"ID": 'PlayButton', "Text": 'Play'}),
                                        ui.Button({"ID": 'LoadButton', "Text": 'Load'}),
                                        ui.Button({"ID": 'ResetButton', "Text": 'Reset'}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        ui.Label({"ID": 'StatusLabel', "Text": ' ', "Alignment": {"AlignHCenter": True, "AlignVCenter": True}}),
                                    ]
                                ),
                                ui.Button(
                                    {
                                        "ID": 'OpenLinkButton',
                                        "Text": '😃Buy Me a Coffee😃，© 2024, Copyright by HB.',
                                        "Alignment": {"AlignHCenter": True, "AlignVCenter": True},
                                        "Font": ui.Font({"PixelSize": 12, "StyleName": 'Bold'}),
                                        "Flat": True,
                                        "TextColor": [0.1, 0.3, 0.9, 1],
                                        "BackgroundColor": [1, 1, 1, 0],
                                        "Weight": 0.1,
                                    }
                                ),
                            ]
                        ),
                        ui.VGroup(
                            [

                            
                                ui.HGroup(
                                    {"Weight": 0.05},
                                    [
                                        ui.Label({"ID": 'RegionLabel', "Text": 'Region', "Alignment": {"AlignRight": False}, "Weight": 0.1}),
                                        ui.LineEdit({"ID": 'Region', "Text": '',  "Weight": 0.2}),
                                        ui.Label({"ID": 'ApiKeyLabel', "Text": 'API Key', "Alignment": {"AlignRight": False}, "Weight": 0.1}),
                                        ui.LineEdit({"ID": 'ApiKey', "Text": '', "EchoMode": 'Password', "Weight": 0.6}),

                                        
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.05},
                                    [
                                        ui.Label({"ID": 'PathLabel', "Text": 'Save Path', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Button({"ID": 'Browse', "Text": 'Browse', "Weight": 0.2}),
                                        ui.LineEdit({"ID": 'Path', "Text": '', "PlaceholderText": '', "ReadOnly": False, "Weight": 0.6}),
                                        
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.05},
                                    [
                                          ui.Label({"ID": 'BalanceLabel', "Text": '',  "Alignment" : {"AlignHCenter" : True, "AlignVCenter" : True},"Weight": 0.2}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.85},
                                    [
                                        ui.TextEdit({"ID": 'infoTxt', "Text": infomsg, "ReadOnly": True}),
                                    ]
                                ),
                                ui.Button(
                                    {
                                        "ID": 'OpenLinkButton',
                                        "Text": '😃Buy Me a Coffee😃，© 2024, Copyright by HB.',
                                        "Alignment": {"AlignHCenter": True, "AlignVCenter": True},
                                        "Font": ui.Font({"PixelSize": 12, "StyleName": 'Bold'}),
                                        "Flat": True,
                                        "TextColor": [0.1, 0.3, 0.9, 1],
                                        "BackgroundColor": [1, 1, 1, 0],
                                        "Weight": 0.1,
                                    }
                                ),
                            ]
                        ),
                    ]
                ),
            ]
        ),
    ]
)

itm = win.GetItems()
# 汉化界面
itm["GetSubButton"].Text = "从时间线获取字幕"
itm["BreakButton"].Text = "停顿"
itm["LanguageLabel"].Text = "语言"
itm["MultilingualLabel"].Text = "语言技能"
itm["NameTypeLabel"].Text = "类型"
itm["NameLabel"].Text = "名称"
itm["StyleLabel"].Text = "风格"
itm["StyleDegreeLabel"].Text = "风格强度"
itm["RateLabel"].Text = "语速"
itm["PitchLabel"].Text = "音高"
itm["VolumeLabel"].Text = "音量"
itm["OutputFormatLabel"].Text = "输出格式"
itm["PlayButton"].Text = "播放预览"
itm["LoadButton"].Text = "合成并加载"
itm["ResetButton"].Text = "重置"
itm["RegionLabel"].Text = "区域"
itm["ApiKeyLabel"].Text = "API 密钥"
itm["PathLabel"].Text = "保存路径"
itm["Browse"].Text = "浏览"
itm["StatusLabel"].Text = ""
itm["MyStack"].CurrentIndex = 0
itm["MyTabs"].AddTab("文字转语音")
itm["MyTabs"].AddTab("配置")

def on_style_degree_slider_value_changed(ev):
    value = ev['Value'] / 100.0  
    itm["StyleDegreeSpinBox"].Value = value
win.On.StyleDegreeSlider.ValueChanged = on_style_degree_slider_value_changed

def on_style_degree_spinbox_value_changed(ev):
    value = int(ev['Value'] * 100)  
    itm["StyleDegreeSlider"].Value = value
win.On.StyleDegreeSpinBox.ValueChanged = on_style_degree_spinbox_value_changed


def on_rate_slider_value_changed(ev):
    value = ev['Value'] / 100.0  
    itm["RateSpinBox"].Value = value
win.On.RateSlider.ValueChanged = on_rate_slider_value_changed

def on_rate_spinbox_value_changed(ev):
    value = int(ev['Value'] * 100)  
    itm["RateSlider"].Value = value
win.On.RateSpinBox.ValueChanged = on_rate_spinbox_value_changed

def on_pitch_slider_value_changed(ev):
    value = ev['Value'] / 100.0  
    itm["PitchSpinBox"].Value = value
win.On.PitchSlider.ValueChanged = on_pitch_slider_value_changed

def on_pitch_spinbox_value_changed(ev):
    value = int(ev['Value'] * 100)  
    itm["PitchSlider"].Value = value
win.On.PitchSpinBox.ValueChanged = on_pitch_spinbox_value_changed

def on_volume_slider_value_changed(ev):
    value = ev['Value'] / 100.0  
    itm["VolumeSpinBox"].Value = value
win.On.VolumeSlider.ValueChanged = on_volume_slider_value_changed

def on_volume_spinbox_value_changed(ev):
    value = int(ev['Value'] * 100)  
    itm["VolumeSlider"].Value = value
win.On.VolumeSpinBox.ValueChanged = on_volume_spinbox_value_changed

def on_my_tabs_current_changed(ev):
    itm["MyStack"].CurrentIndex = ev["Index"]
win.On.MyTabs.CurrentChanged = on_my_tabs_current_changed
audio_formats = {
    "8k, .wav": speechsdk.SpeechSynthesisOutputFormat.Riff8Khz16BitMonoPcm,
    "16k, .wav": speechsdk.SpeechSynthesisOutputFormat.Riff16Khz16BitMonoPcm,
    "24k, .wav": speechsdk.SpeechSynthesisOutputFormat.Riff24Khz16BitMonoPcm,
    "48k, .wav": speechsdk.SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm,
   # "16k, .mp3": speechsdk.SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3,
   # "24k, .mp3": speechsdk.SpeechSynthesisOutputFormat.Audio24Khz48KBitRateMonoMp3,
   # "48k, .mp3": speechsdk.SpeechSynthesisOutputFormat.Audio48Khz96KBitRateMonoMp3
}
# 添加音频格式选项到 OutputFormatCombo
for fmt in audio_formats.keys():
    itm["OutputFormatCombo"].AddItem(fmt)

# 初始化 subtitle 变量
subtitle = ""
lang = ""
multilingual='默认'
current_context=""
ssml = ''
voice_name = ""
style = None
rate = None
pitch = None
volume = None
style_degree = None
stream = None

def on_subtitle_text_changed(ev):
    global stream
    stream = None
win.On.SubtitleTxt.TextChanged = on_subtitle_text_changed

voices = {
    'zh-CN': {
        #'language': 'Chinese (Simplified)',
        'language': 'Chinese (普通话)',
        'voices': [
            "zh-CN-XiaoxiaoNeural (Female)",
            "zh-CN-YunxiNeural (Male)",
            "zh-CN-YunjianNeural (Male)",
            "zh-CN-XiaoyiNeural (Female)",
            "zh-CN-YunyangNeural (Male)",
            "zh-CN-XiaochenNeural (Female)",
            "zh-CN-XiaohanNeural (Female)",
            "zh-CN-XiaomengNeural (Female)",
            "zh-CN-XiaomoNeural (Female)",
            "zh-CN-XiaoqiuNeural (Female)",
            "zh-CN-XiaoruiNeural (Female)",
            "zh-CN-XiaoshuangNeural (Female, Child)",
            "zh-CN-XiaoyanNeural (Female)",
            "zh-CN-XiaoyouNeural (Female, Child)",
            "zh-CN-XiaozhenNeural (Female)",
            "zh-CN-YunfengNeural (Male)",
            "zh-CN-YunhaoNeural (Male)",
            "zh-CN-YunxiaNeural (Male)",
            "zh-CN-YunyeNeural (Male)",
            "zh-CN-YunzeNeural (Male)",
            "zh-CN-XiaochenMultilingualNeural (Female)",
            "zh-CN-XiaorouNeural (Female)",
            "zh-CN-XiaoxiaoDialectsNeural (Female)",
            "zh-CN-XiaoxiaoMultilingualNeural (Female)",
            "zh-CN-XiaoyuMultilingual (Female)",
            "zh-CN-YunjieNeural (Male)",
            "zh-CN-YunyiMultilingualNeural (Male)"
        ]
    },
    'yue-CN': {
        #'language': 'Chinese (Cantonese, Simplified)',
        'language': 'Chinese (粤语，简体)',
        'voices': [
            "yue-CN-XiaoMinNeural (Female)",
            "yue-CN-YunSongNeural (Male)"
        ]
    },
    'zh-CN-GUANGXI': {
        #'language': 'Chinese (Guangxi Accent Mandarin, Simplified)',
        'language': 'Chinese (广西口音普通话，简体)',
        'voices': [
            "zh-CN-guangxi-YunqiNeural (Male)"
        ]
    },
    'zh-CN-henan': {
        #'language': 'Chinese (Zhongyuan Mandarin Henan, Simplified)',
        'language': 'Chinese (中原官话河南，简体)',
        'voices': [
            "zh-CN-henan-YundengNeural (Male)"
        ]
    },
    'zh-CN-liaoning': {
        #'language': 'Chinese (Northeastern Mandarin, Simplified)',
        'language': 'Chinese (东北官话，简体)',
        'voices': [
            "zh-CN-liaoning-XiaobeiNeural (Female)",
            "zh-CN-liaoning-YunbiaoNeural (Male)"
        ]
    },
    'zh-CN-shaanxi': {
        #'language': 'Chinese (Zhongyuan Mandarin Shaanxi, Simplified)',
        'language': 'Chinese (中原官话陕西，简体)',
        'voices': [
            "zh-CN-shaanxi-XiaoniNeural (Female)"
        ]
    },
    'zh-CN-shandong': {
        #'language': 'Chinese (Jilu Mandarin, Simplified)',
        'language': 'Chinese (冀鲁官话，简体)',
        'voices': [
            "zh-CN-shandong-YunxiangNeural (Male)"
        ]
    },
    'zh-CN-sichuan': {
        #'language': 'Chinese (Southwestern Mandarin, Simplified)',
        'language': 'Chinese (西南官话，简体)',
        'voices': [
            "zh-CN-sichuan-YunxiNeural (Male)"
        ]
    },
    'wuu-CN': {
        #'language': 'Chinese (Wu, Simplified)',
        'language': 'Chinese (吴语，简体)',
        'voices': [
            "wuu-CN-XiaotongNeural (Female)",
            "wuu-CN-YunzheNeural (Male)"
        ]
    },
    'zh-HK': {
        #'language': 'Chinese (Cantonese, Traditional)',
        'language': 'Chinese (粤语，繁体)',
        'voices': [
            "zh-HK-HiuMaanNeural (Female)",
            "zh-HK-WanLungNeural (Male)",
            "zh-HK-HiuGaaiNeural (Female)"
        ]
    },
    'zh-TW': {
        #'language': 'Chinese (Taiwanese Mandarin, Traditional)',
        'language': 'Chinese (台湾，繁体)',
        'voices': [
            "zh-TW-HsiaoChenNeural (Female)",
            "zh-TW-YunJheNeural (Male)",
            "zh-TW-HsiaoYuNeural (Female)"
        ]
    },
    'en-AU': {
        'language': 'English (Australia)',
        'voices': [
            "en-AU-NatashaNeural (Female)",
            "en-AU-WilliamNeural (Male)",
            "en-AU-AnnetteNeural (Female)",
            "en-AU-CarlyNeural (Female)",
            "en-AU-DarrenNeural (Male)",
            "en-AU-DuncanNeural (Male)",
            "en-AU-ElsieNeural (Female)",
            "en-AU-FreyaNeural (Female)",
            "en-AU-JoanneNeural (Female)",
            "en-AU-KenNeural (Male)",
            "en-AU-KimNeural (Female)",
            "en-AU-NeilNeural (Male)",
            "en-AU-TimNeural (Male)",
            "en-AU-TinaNeural (Female)"
        ]
    },
    'en-CA': {
        'language': 'English (Canada)',
        'voices': [
            "en-CA-ClaraNeural (Female)",
            "en-CA-LiamNeural (Male)"
        ]
    },
    'en-GB': {
        'language': 'English (United Kingdom)',
        'voices': [
            "en-GB-SoniaNeural (Female)",
            "en-GB-RyanNeural (Male)",
            "en-GB-LibbyNeural (Female)",
            "en-GB-AbbiNeural (Female)",
            "en-GB-AdaMultilingualNeural (Female)",
            "en-GB-AlfieNeural (Male)",
            "en-GB-BellaNeural (Female)",
            "en-GB-ElliotNeural (Male)",
            "en-GB-EthanNeural (Male)",
            "en-GB-HollieNeural (Female)",
            "en-GB-MaisieNeural (Female, Child)",
            "en-GB-NoahNeural (Male)",
            "en-GB-OliverNeural (Male)",
            "en-GB-OliviaNeural (Female)",
            "en-GB-ThomasNeural (Male)",
            "en-GB-OllieMultilingualNeural (Male)"
        ]
    },
    'en-HK': {
        'language': 'English (Hong Kong SAR)',
        'voices': [
            "en-HK-YanNeural (Female)",
            "en-HK-SamNeural (Male)"
        ]
    },
    'en-IE': {
        'language': 'English (Ireland)',
        'voices': [
            "en-IE-EmilyNeural (Female)",
            "en-IE-ConnorNeural (Male)"
        ]
    },
    'en-IN': {
        'language': 'English (India)',
        'voices': [
            "en-IN-NeerjaNeural (Female)",
            "en-IN-PrabhatNeural (Male)"
        ]
    },
    'en-KE': {
        'language': 'English (Kenya)',
        'voices': [
            "en-KE-AsiliaNeural (Female)",
            "en-KE-ChilembaNeural (Male)"
        ]
    },
    'en-NG': {
        'language': 'English (Nigeria)',
        'voices': [
            "en-NG-EzinneNeural (Female)",
            "en-NG-AbeoNeural (Male)"
        ]
    },
    'en-NZ': {
        'language': 'English (New Zealand)',
        'voices': [
            "en-NZ-MollyNeural (Female)",
            "en-NZ-MitchellNeural (Male)"
        ]
    },
    'en-PH': {
        'language': 'English (Philippines)',
        'voices': [
            "en-PH-RosaNeural (Female)",
            "en-PH-JamesNeural (Male)"
        ]
    },
    'en-SG': {
        'language': 'English (Singapore)',
        'voices': [
            "en-SG-LunaNeural (Female)",
            "en-SG-WayneNeural (Male)"
        ]
    },
    'en-TZ': {
        'language': 'English (Tanzania)',
        'voices': [
            "en-TZ-ImaniNeural (Female)",
            "en-TZ-ElimuNeural (Male)"
        ]
    },
    'en-US': {
        'language': 'English (United States)',
        'voices': [
            "en-US-AvaNeural (Female)",
            "en-US-AndrewNeural (Male)",
            "en-US-EmmaNeural (Female)",
            "en-US-BrianNeural (Male)",
            "en-US-JennyNeural (Female)",
            "en-US-GuyNeural (Male)",
            "en-US-AriaNeural (Female)",
            "en-US-DavisNeural (Male)",
            "en-US-JaneNeural (Female)",
            "en-US-JasonNeural (Male)",
            "en-US-SaraNeural (Female)",
            "en-US-TonyNeural (Male)",
            "en-US-NancyNeural (Female)",
            "en-US-AmberNeural (Female)",
            "en-US-AnaNeural (Female, Child)",
            "en-US-AshleyNeural (Female)",
            "en-US-BrandonNeural (Male)",
            "en-US-ChristopherNeural (Male)",
            "en-US-CoraNeural (Female)",
            "en-US-ElizabethNeural (Female)",
            "en-US-EricNeural (Male)",
            "en-US-JacobNeural (Male)",
            "en-US-JennyMultilingualNeural (Female)",
            "en-US-MichelleNeural (Female)",
            "en-US-MonicaNeural (Female)",
            "en-US-RogerNeural (Male)",
            "en-US-RyanMultilingualNeural (Male)",
            "en-US-SteffanNeural (Male)",
            "en-US-AIGenerate1Neural1 (Male)",
            "en-US-AIGenerate2Neural1 (Female)",
            "en-US-AndrewMultilingualNeural (Male)",
            "en-US-AvaMultilingualNeural (Female)",
            "en-US-BlueNeural (Neutral)",
            "en-US-KaiNeural (Male)",
            "en-US-LunaNeural (Female)",
            "en-US-BrianMultilingualNeural (Male)",
            "en-US-EmmaMultilingualNeural (Female)",
            "en-US-AlloyMultilingualNeural (Male)",
            "en-US-EchoMultilingualNeural (Male)",
            "en-US-FableMultilingualNeural (Neutral)",
            "en-US-OnyxMultilingualNeural (Male)",
            "en-US-NovaMultilingualNeural (Female)",
            "en-US-ShimmerMultilingualNeural (Female)",
            "en-US-AlloyMultilingualNeuralHD (Male)",
            "en-US-EchoMultilingualNeuralHD (Male)",
            "en-US-FableMultilingualNeuralHD (Neutral)",
            "en-US-OnyxMultilingualNeuralHD (Male)",
            "en-US-NovaMultilingualNeuralHD (Female)",
            "en-US-ShimmerMultilingualNeuralHD (Female)"
        ]
    },
    'es-AR': {
        'language': 'Spanish (Argentina)',
        'voices': [
            "es-AR-ElenaNeural (Female)",
            "es-AR-TomasNeural (Male)"
        ]
    },
    'es-BO': {
        'language': 'Spanish (Bolivia)',
        'voices': [
            "es-BO-SofiaNeural (Female)",
            "es-BO-MarceloNeural (Male)"
        ]
    },
    'es-CL': {
        'language': 'Spanish (Chile)',
        'voices': [
            "es-CL-CatalinaNeural (Female)",
            "es-CL-LorenzoNeural (Male)"
        ]
    },
    'es-CO': {
        'language': 'Spanish (Colombia)',
        'voices': [
            "es-CO-SalomeNeural (Female)",
            "es-CO-GonzaloNeural (Male)"
        ]
    },
    'es-CR': {
        'language': 'Spanish (Costa Rica)',
        'voices': [
            "es-CR-MariaNeural (Female)",
            "es-CR-JuanNeural (Male)"
        ]
    },
    'es-CU': {
        'language': 'Spanish (Cuba)',
        'voices': [
            "es-CU-BelkysNeural (Female)",
            "es-CU-ManuelNeural (Male)"
        ]
    },
    'es-DO': {
        'language': 'Spanish (Dominican Republic)',
        'voices': [
            "es-DO-RamonaNeural (Female)",
            "es-DO-EmilioNeural (Male)"
        ]
    },
    'es-EC': {
        'language': 'Spanish (Ecuador)',
        'voices': [
            "es-EC-AndreaNeural (Female)",
            "es-EC-LuisNeural (Male)"
        ]
    },
    'es-ES': {
        'language': 'Spanish (Spain)',
        'voices': [
            "es-ES-ElviraNeural (Female)",
            "es-ES-AlvaroNeural (Male)",
            "es-ES-AbrilNeural (Female)",
            "es-ES-ArnauNeural (Male)",
            "es-ES-DarioNeural (Male)",
            "es-ES-EliasNeural (Male)",
            "es-ES-EstrellaNeural (Female)",
            "es-ES-IreneNeural (Female)",
            "es-ES-LaiaNeural (Female)",
            "es-ES-LiaNeural (Female)",
            "es-ES-NilNeural (Male)",
            "es-ES-SaulNeural (Male)",
            "es-ES-TeoNeural (Male)",
            "es-ES-TrianaNeural (Female)",
            "es-ES-VeraNeural (Female)",
            "es-ES-XimenaNeural (Female)",
            "es-ES-ArabellaMultilingualNeural (Female)",
            "es-ES-IsidoraMultilingualNeural (Female)"
        ]
    },
    'es-GQ': {
        'language': 'Spanish (Equatorial Guinea)',
        'voices': [
            "es-GQ-TeresaNeural (Female)",
            "es-GQ-JavierNeural (Male)"
        ]
    },
    'es-GT': {
        'language': 'Spanish (Guatemala)',
        'voices': [
            "es-GT-MartaNeural (Female)",
            "es-GT-AndresNeural (Male)"
        ]
    },
    'es-HN': {
        'language': 'Spanish (Honduras)',
        'voices': [
            "es-HN-KarlaNeural (Female)",
            "es-HN-CarlosNeural (Male)"
        ]
    },
    'es-MX': {
        'language': 'Spanish (Mexico)',
        'voices': [
            "es-MX-DaliaNeural (Female)",
            "es-MX-JorgeNeural (Male)",
            "es-MX-BeatrizNeural (Female)",
            "es-MX-CandelaNeural (Female)",
            "es-MX-CarlotaNeural (Female)",
            "es-MX-CecilioNeural (Male)",
            "es-MX-GerardoNeural (Male)",
            "es-MX-LarissaNeural (Female)",
            "es-MX-LibertoNeural (Male)",
            "es-MX-LucianoNeural (Male)",
            "es-MX-MarinaNeural (Female)",
            "es-MX-NuriaNeural (Female)",
            "es-MX-PelayoNeural (Male)",
            "es-MX-RenataNeural (Female)",
            "es-MX-YagoNeural (Male)"
        ]
    },
    'es-NI': {
        'language': 'Spanish (Nicaragua)',
        'voices': [
            "es-NI-YolandaNeural (Female)",
            "es-NI-FedericoNeural (Male)"
        ]
    },
    'es-PA': {
        'language': 'Spanish (Panama)',
        'voices': [
            "es-PA-MargaritaNeural (Female)",
            "es-PA-RobertoNeural (Male)"
        ]
    },
    'es-PE': {
        'language': 'Spanish (Peru)',
        'voices': [
            "es-PE-CamilaNeural (Female)",
            "es-PE-AlexNeural (Male)"
        ]
    },
    'es-PR': {
        'language': 'Spanish (Puerto Rico)',
        'voices': [
            "es-PR-KarinaNeural (Female)",
            "es-PR-VictorNeural (Male)"
        ]
    },
    'es-PY': {
        'language': 'Spanish (Paraguay)',
        'voices': [
            "es-PY-TaniaNeural (Female)",
            "es-PY-MarioNeural (Male)"
        ]
    },
    'es-SV': {
        'language': 'Spanish (El Salvador)',
        'voices': [
            "es-SV-LorenaNeural (Female)",
            "es-SV-RodrigoNeural (Male)"
        ]
    },
    'es-US': {
        'language': 'Spanish (United States)',
        'voices': [
            "es-US-PalomaNeural (Female)",
            "es-US-AlonsoNeural (Male)"
        ]
    },
    'es-UY': {
        'language': 'Spanish (Uruguay)',
        'voices': [
            "es-UY-ValentinaNeural (Female)",
            "es-UY-MateoNeural (Male)"
        ]
    },
    'es-VE': {
        'language': 'Spanish (Venezuela)',
        'voices': [
            "es-VE-PaolaNeural (Female)",
            "es-VE-SebastianNeural (Male)"
        ]
    },
    'de-AT': {
        'language': 'German (Austria)',
        'voices': [
            "de-AT-IngridNeural (Female)",
            "de-AT-JonasNeural (Male)"
        ]
    },
    'de-CH': {
        'language': 'German (Switzerland)',
        'voices': [
            "de-CH-LeniNeural (Female)",
            "de-CH-JanNeural (Male)"
        ]
    },
    'de-DE': {
        'language': 'German (Germany)',
        'voices': [
            "de-DE-KatjaNeural (Female)",
            "de-DE-ConradNeural (Male)",
            "de-DE-AmalaNeural (Female)",
            "de-DE-BerndNeural (Male)",
            "de-DE-ChristophNeural (Male)",
            "de-DE-ElkeNeural (Female)",
            "de-DE-GiselaNeural (Female, Child)",
            "de-DE-KasperNeural (Male)",
            "de-DE-KillianNeural (Male)",
            "de-DE-KlarissaNeural (Female)",
            "de-DE-KlausNeural (Male)",
            "de-DE-LouisaNeural (Female)",
            "de-DE-MajaNeural (Female)",
            "de-DE-RalfNeural (Male)",
            "de-DE-TanjaNeural (Female)",
            "de-DE-FlorianMultilingualNeural (Male)",
            "de-DE-SeraphinaMultilingualNeural (Female)"
        ]
    },
    'fr-BE': {
        'language': 'French (Belgium)',
        'voices': [
            "fr-BE-CharlineNeural (Female)",
            "fr-BE-GerardNeural (Male)"
        ]
    },
    'fr-CA': {
        'language': 'French (Canada)',
        'voices': [
            "fr-CA-SylvieNeural (Female)",
            "fr-CA-JeanNeural (Male)",
            "fr-CA-AntoineNeural (Male)",
            "fr-CA-ThierryNeural1 (Male)"
        ]
    },
    'fr-CH': {
        'language': 'French (Switzerland)',
        'voices': [
            "fr-CH-ArianeNeural (Female)",
            "fr-CH-FabriceNeural (Male)"
        ]
    },
    'fr-FR': {
        'language': 'French (France)',
        'voices': [
            "fr-FR-DeniseNeural (Female)",
            "fr-FR-HenriNeural (Male)",
            "fr-FR-AlainNeural (Male)",
            "fr-FR-BrigitteNeural (Female)",
            "fr-FR-CelesteNeural (Female)",
            "fr-FR-ClaudeNeural (Male)",
            "fr-FR-CoralieNeural (Female)",
            "fr-FR-EloiseNeural (Female, Child)",
            "fr-FR-JacquelineNeural (Female)",
            "fr-FR-JeromeNeural (Male)",
            "fr-FR-JosephineNeural (Female)",
            "fr-FR-MauriceNeural (Male)",
            "fr-FR-YvesNeural (Male)",
            "fr-FR-YvetteNeural (Female)",
            "fr-FR-RemyMultilingualNeural (Male)",
            "fr-FR-VivienneMultilingualNeural (Female)"
        ]
    },
    'it-IT': {
        'language': 'Italian (Italy)',
        'voices': [
            "it-IT-ElsaNeural (Female)",
            "it-IT-IsabellaNeural (Female)",
            "it-IT-DiegoNeural (Male)",
            "it-IT-BenignoNeural (Male)",
            "it-IT-CalimeroNeural (Male)",
            "it-IT-CataldoNeural (Male)",
            "it-IT-FabiolaNeural (Female)",
            "it-IT-FiammaNeural (Female)",
            "it-IT-GianniNeural (Male)",
            "it-IT-ImeldaNeural (Female)",
            "it-IT-IrmaNeural (Female)",
            "it-IT-LisandroNeural (Male)",
            "it-IT-PalmiraNeural (Female)",
            "it-IT-PierinaNeural (Female)",
            "it-IT-RinaldoNeural (Male)",
            "it-IT-AlessioMultilingualNeural (Male)",
            "it-IT-IsabellaMultilingualNeural (Female)",
            "it-IT-MarcelloMultilingualNeural (Male)",
            "it-IT-GiuseppeNeural (Male)"
        ]
    },
    'ja-JP': {
        'language': 'Japanese (Japan)',
        'voices': [
            "ja-JP-NanamiNeural (Female)",
            "ja-JP-KeitaNeural (Male)",
            "ja-JP-AoiNeural (Female)",
            "ja-JP-DaichiNeural (Male)",
            "ja-JP-MayuNeural (Female)",
            "ja-JP-NaokiNeural (Male)",
            "ja-JP-ShioriNeural (Female)",
            "ja-JP-MasaruMultilingualNeural (Male)"
        ]
    },
    'ko-KR': {
        'language': 'Korean (Korea)',
        'voices': [
            "ko-KR-SunHiNeural (Female)",
            "ko-KR-InJoonNeural (Male)",
            "ko-KR-BongJinNeural (Male)",
            "ko-KR-GookMinNeural (Male)",
            "ko-KR-JiMinNeural (Female)",
            "ko-KR-SeoHyeonNeural (Female)",
            "ko-KR-SoonBokNeural (Female)",
            "ko-KR-YuJinNeural (Female)",
            "ko-KR-HyunsuNeural1 (Male)"
        ]
    },
    'zu-ZA': {
        'language': 'Zulu (South Africa)',
        'voices': [
            "zu-ZA-ThandoNeural (Female)",
            "zu-ZA-ThembaNeural (Male)"
        ]
    },
}

styles = {
    "de-DE-ConradNeural": ["cheerful"],
    "en-US-AriaNeural": ["angry", "chat", "cheerful", "customerservice", "empathetic", "excited", "friendly", "hopeful", "narration-professional", "newscast-casual", "newscast-formal", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-DavisNeural": ["angry", "chat", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-GuyNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "newscast", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-JaneNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-JasonNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-JennyNeural": ["angry", "assistant", "chat", "cheerful", "customerservice", "excited", "friendly", "hopeful", "newscast", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-NancyNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-SaraNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "en-US-TonyNeural": ["angry", "cheerful", "excited", "friendly", "hopeful", "sad", "shouting", "terrified", "unfriendly", "whispering"],
    "zh-CN-XiaohanNeural": ["affectionate", "angry", "calm", "cheerful", "disgruntled", "embarrassed", "fearful", "gentle", "sad", "serious"],
    "zh-CN-XiaomengNeural": ["chat"],
    "zh-CN-XiaomoNeural": ["affectionate", "angry", "calm", "cheerful", "depressed", "disgruntled", "embarrassed", "envious", "fearful", "gentle", "sad", "serious"],
    "zh-CN-XiaoruiNeural": ["angry", "calm", "fearful", "sad"],
    "zh-CN-XiaoshuangNeural": ["chat"],
    "zh-CN-XiaoxiaoNeural": ["affectionate", "angry", "assistant", "calm", "chat", "chat-casual", "cheerful", "customerservice", "disgruntled", "fearful", "friendly", "gentle", "lyrical", "newscast", "poetry-reading", "sad", "serious", "sorry", "whisper"],
    "zh-CN-XiaoyiNeural": ["affectionate", "angry", "cheerful", "disgruntled", "embarrassed", "fearful", "gentle", "sad", "serious"],
    "zh-CN-XiaozhenNeural": ["angry", "cheerful", "disgruntled", "fearful", "sad", "serious"],
    "zh-CN-YunfengNeural": ["angry", "cheerful", "depressed", "disgruntled", "fearful", "sad", "serious"],
    "zh-CN-YunhaoNeural": ["advertisement-upbeat"],
    "zh-CN-YunjianNeural": ["angry", "cheerful", "depressed", "disgruntled", "documentary-narration", "narration-relaxed", "sad", "serious", "sports-commentary", "sports-commentary-excited"],
    "zh-CN-YunxiaNeural": ["angry", "calm", "cheerful", "fearful", "sad"],
    "zh-CN-YunxiNeural": ["angry", "assistant", "chat", "cheerful", "depressed", "disgruntled", "embarrassed", "fearful", "narration-relaxed", "newscast", "sad", "serious"],
    "zh-CN-YunyangNeural": ["customerservice", "narration-professional", "newscast-casual"],
    "zh-CN-YunyeNeural": ["angry", "calm", "cheerful", "disgruntled", "embarrassed", "fearful", "sad", "serious"],
    "zh-CN-YunzeNeural": ["angry", "calm", "cheerful", "depressed", "disgruntled", "documentary-narration", "fearful", "sad", "serious"]
}
# 汉化映射字典
style_translation = {
    "cheerful": "高兴",
    "angry": "愤怒",
    "chat": "聊天",
    "customerservice": "客服",
    "empathetic": "有同情心的",
    "excited": "兴奋",
    "friendly": "友好",
    "hopeful": "有希望的",
    "narration-professional": "专业叙述",
    "newscast-casual": "新闻播报-休闲",
    "newscast-formal": "新闻播报-正式",
    "newscast": "新闻播报",
    "sad": "悲伤",
    "shouting": "喊叫",
    "terrified": "害怕",
    "unfriendly": "不友好",
    "whispering": "耳语",
    "whisper": "耳语",
    "affectionate": "亲切",
    "calm": "平静",
    "disgruntled": "不满",
    "embarrassed": "尴尬",
    "fearful": "害怕",
    "gentle": "温柔",
    "serious": "严肃",
    "assistant": "助手",
    "chat-casual": "聊天-休闲",
    "lyrical": "抒情",
    "poetry-reading": "诗歌朗诵",
    "sorry": "抱歉",
    "advertisement-upbeat": "广告-积极",
    "depressed": "沮丧",
    "envious": "嫉妒",
    "documentary-narration": "纪录片叙述",
    "narration-relaxed": "叙述-放松",
    "sports-commentary": "体育评论",
    "sports-commentary-excited": "体育评论-兴奋"
}
# 翻译成中文
def translate_styles_to_chinese(style_code):
    return style_translation.get(style_code, style_code)

def get_original_style(chinese_style):
    for eng, chi in style_translation.items():
        if chi == chinese_style:
            return eng
    return chinese_style  # 如果找不到对应的英文风格，返回原始中文风格

# 定义多语言和人名组
Multilinguals = {
    "Multilingual1": {
        "names": [
            "en-US-AndrewMultilingualNeural (Male)", "en-US-AvaMultilingualNeural (Female)", "en-US-BrianMultilingualNeural (Male)", 
            "en-US-EmmaMultilingualNeural (Female)", "en-GB-AdaMultilingualNeural (Female)", "en-GB-OllieMultilingualNeural (Male)", 
            "de-DE-SeraphinaMultilingualNeural (Female)", "de-DE-FlorianMultilingualNeural (Male)", "es-ES-IsidoraMultilingualNeural (Female)", 
            "es-ES-ArabellaMultilingualNeural (Female)", "fr-FR-VivienneMultilingualNeural (Female)", "fr-FR-RemyMultilingualNeural (Male)", 
            "it-IT-IsabellaMultilingualNeural (Female)", "it-IT-MarcelloMultilingualNeural (Male)", "it-IT-AlessioMultilingualNeural (Male)", 
            "ja-JP-MasaruMultilingualNeural (Male)", "pt-BR-ThalitaMultilingualNeural (Female)", "zh-CN-XiaoxiaoMultilingualNeural (Female)", 
            "zh-CN-XiaochenMultilingualNeural (Female)", "zh-CN-YunyiMultilingualNeural (Male)"
        ],
        "languages": {
            "af-ZA", "sq-AL", "am-ET", "ar-EG", "ar-SA", "hy-AM", "az-AZ", "eu-ES", "bn-IN", "bs-BA", "bg-BG", "my-MM", "ca-ES", "zh-HK", 
            "zh-CN", "zh-TW", "hr-HR", "cs-CZ", "da-DK", "nl-BE", "nl-NL", "en-AU", "en-CA", "en-HK", "en-IN", "en-IE", "en-GB", "en-US", 
            "et-EE", "fil-PH", "fi-FI", "fr-BE", "fr-CA", "fr-FR", "fr-CH", "gl-ES", "ka-GE", "de-AT", "de-DE", "de-CH", "el-GR", "he-IL", 
            "hi-IN", "hu-HU", "is-IS", "id-ID", "ga-IE", "it-IT", "ja-JP", "jv-ID", "kn-IN", "kk-KZ", "km-KH", "ko-KR", "lo-LA", "lv-LV", 
            "lt-LT", "mk-MK", "ms-MY", "ml-IN", "mt-MT", "mn-MN", "ne-NP", "nb-NO", "ps-AF", "fa-IR", "pl-PL", "pt-BR", "pt-PT", "ro-RO", 
            "ru-RU", "sr-RS", "si-LK", "sk-SK", "sl-SI", "so-SO", "es-MX", "es-ES", "su-ID", "sw-KE", "sv-SE", "ta-IN", "te-IN", "th-TH", 
            "tr-TR", "uk-UA", "ur-PK", "uz-UZ", "vi-VN", "cy-GB", "zu-ZA"
        }
    },
    "Multilingual2": {
        "names": [
            "en-US-AlloyMultilingualNeural (Male)", "en-US-EchoMultilingualNeural (Male)", "en-US-FableMultilingualNeural (Neutral)", 
            "en-US-OnyxMultilingualNeural (Male)", "en-US-NovaMultilingualNeural (Female)", "en-US-ShimmerMultilingualNeural (Female)", 
            "en-US-AlloyMultilingualNeuralHD (Male)", "en-US-EchoMultilingualNeuralHD (Male)", "en-US-FableMultilingualNeuralHD (Neutral)", 
            "en-US-OnyxMultilingualNeuralHD (Male)", "en-US-NovaMultilingualNeuralHD (Female)", "en-US-ShimmerMultilingualNeuralHD (Female)"
        ],
        "languages": {
            "af-ZA", "ar-EG", "hy-AM", "az-AZ", "be-BY", "bs-BA", "bg-BG", "ca-ES", "zh-CN", "hr-HR", "cs-CZ", "da-DK", "nl-NL", "en-US", 
            "et-EE", "fi-FI", "fr-FR", "gl-ES", "de-DE", "el-GR", "he-IL", "hi-IN", "hu-HU", "is-IS", "id-ID", "it-IT", "ja-JP", "kn-IN", 
            "kk-KZ", "ko-KR", "lv-LV", "lt-LT", "mk-MK", "ms-MY", "mr-IN", "mi-NZ", "ne-NP", "nb-NO", "fa-IR", "pl-PL", "pt-BR", "ro-RO", 
            "ru-RU", "sr-RS", "sk-SK", "sl-SI", "es-ES", "sw-KE", "sv-SE", "fil-PH", "ta-IN", "th-TH", "tr-TR", "uk-UA", "ur-PK", "vi-VN", 
            "cy-GB"
        }
    },
    "Multilingual3": {
        "names": ["en-US-JennyMultilingualNeural (Female)", "en-US-RyanMultilingualNeural (Male)"],
        "languages": {
            "ar-EG", "ar-SA", "ca-ES", "zh-HK", "zh-CN", "zh-TW", "cs-CZ", "da-DK", "nl-BE", "nl-NL", "en-AU", "en-CA", "en-HK", "en-IN", 
            "en-IE", "en-GB", "en-US", "fi-FI", "fr-BE", "fr-CA", "fr-FR", "fr-CH", "de-AT", "de-DE", "de-CH", "hi-IN", "hu-HU", "id-ID", 
            "it-IT", "ja-JP", "ko-KR", "nb-NO", "pl-PL", "pt-BR", "pt-PT", "ru-RU", "es-MX", "es-ES", "sv-SE", "th-TH", "tr-TR"
        }
    }
}

# 汉化映射字典
lang_translation = {
    "af-ZA": "南非语（南非）",
    "sq-AL": "阿尔巴尼亚语（阿尔巴尼亚）",
    "am-ET": "阿姆哈拉语（埃塞俄比亚）",
    "ar-EG": "阿拉伯语（埃及）",
    "ar-SA": "阿拉伯语（沙特阿拉伯）",
    "hy-AM": "亚美尼亚语（亚美尼亚）",
    "az-AZ": "阿塞拜疆语（阿塞拜疆）",
    "eu-ES": "巴斯克语（西班牙）",
    "bn-IN": "孟加拉语（印度）",
    "bs-BA": "波斯尼亚语（波斯尼亚和黑塞哥维那）",
    "bg-BG": "保加利亚语（保加利亚）",
    "my-MM": "缅甸语（缅甸）",
    "ca-ES": "加泰罗尼亚语（西班牙）",
    "zh-HK": "中文（香港）",
    "zh-CN": "中文（简体，中国）",
    "zh-TW": "中文（繁体，台湾）",
    "hr-HR": "克罗地亚语（克罗地亚）",
    "cs-CZ": "捷克语（捷克）",
    "da-DK": "丹麦语（丹麦）",
    "nl-BE": "荷兰语（比利时）",
    "nl-NL": "荷兰语（荷兰）",
    "en-AU": "英语（澳大利亚）",
    "en-CA": "英语（加拿大）",
    "en-HK": "英语（香港）",
    "en-IN": "英语（印度）",
    "en-IE": "英语（爱尔兰）",
    "en-GB": "英语（英国）",
    "en-US": "英语（美国）",
    "et-EE": "爱沙尼亚语（爱沙尼亚）",
    "fil-PH": "菲律宾语（菲律宾）",
    "fi-FI": "芬兰语（芬兰）",
    "fr-BE": "法语（比利时）",
    "fr-CA": "法语（加拿大）",
    "fr-FR": "法语（法国）",
    "fr-CH": "法语（瑞士）",
    "gl-ES": "加利西亚语（西班牙）",
    "ka-GE": "格鲁吉亚语（格鲁吉亚）",
    "de-AT": "德语（奥地利）",
    "de-DE": "德语（德国）",
    "de-CH": "德语（瑞士）",
    "el-GR": "希腊语（希腊）",
    "he-IL": "希伯来语（以色列）",
    "hi-IN": "印地语（印度）",
    "hu-HU": "匈牙利语（匈牙利）",
    "is-IS": "冰岛语（冰岛）",
    "id-ID": "印度尼西亚语（印度尼西亚）",
    "ga-IE": "爱尔兰语（爱尔兰）",
    "it-IT": "意大利语（意大利）",
    "ja-JP": "日语（日本）",
    "jv-ID": "爪哇语（印度尼西亚）",
    "kn-IN": "卡纳达语（印度）",
    "kk-KZ": "哈萨克语（哈萨克斯坦）",
    "km-KH": "高棉语（柬埔寨）",
    "ko-KR": "韩语（韩国）",
    "lo-LA": "老挝语（老挝）",
    "lv-LV": "拉脱维亚语（拉脱维亚）",
    "lt-LT": "立陶宛语（立陶宛）",
    "mk-MK": "马其顿语（北马其顿）",
    "ms-MY": "马来语（马来西亚）",
    "ml-IN": "马拉雅拉姆语（印度）",
    "mt-MT": "马耳他语（马耳他）",
    "mn-MN": "蒙古语（蒙古）",
    "ne-NP": "尼泊尔语（尼泊尔）",
    "nb-NO": "挪威语（挪威）",
    "ps-AF": "普什图语（阿富汗）",
    "fa-IR": "波斯语（伊朗）",
    "pl-PL": "波兰语（波兰）",
    "pt-BR": "葡萄牙语（巴西）",
    "pt-PT": "葡萄牙语（葡萄牙）",
    "ro-RO": "罗马尼亚语（罗马尼亚）",
    "ru-RU": "俄语（俄罗斯）",
    "sr-RS": "塞尔维亚语（塞尔维亚）",
    "si-LK": "僧伽罗语（斯里兰卡）",
    "sk-SK": "斯洛伐克语（斯洛伐克）",
    "sl-SI": "斯洛文尼亚语（斯洛文尼亚）",
    "so-SO": "索马里语（索马里）",
    "es-MX": "西班牙语（墨西哥）",
    "es-ES": "西班牙语（西班牙）",
    "su-ID": "巽他语（印度尼西亚）",
    "sw-KE": "斯瓦希里语（肯尼亚）",
    "sv-SE": "瑞典语（瑞典）",
    "ta-IN": "泰米尔语（印度）",
    "te-IN": "泰卢固语（印度）",
    "th-TH": "泰语（泰国）",
    "tr-TR": "土耳其语（土耳其）",
    "uk-UA": "乌克兰语（乌克兰）",
    "ur-PK": "乌尔都语（巴基斯坦）",
    "uz-UZ": "乌兹别克语（乌兹别克斯坦）",
    "vi-VN": "越南语（越南）",
    "cy-GB": "威尔士语（英国）",
    "zu-ZA": "祖鲁语（南非）"
}

# 翻译成中文
def translate_lang_to_chinese(lang_code):
    return lang_translation.get(lang_code, lang_code)

def get_original_lang(chinese_lang):
    for lang_code, lang_name in lang_translation.items():
        if lang_name == chinese_lang:
            return lang_code
    return chinese_lang

Language = [voices[locale]['language'] for locale in voices.keys()]
for language in Language:
    itm["LanguageCombo"].AddItem(language)

NameType = ['Female','Male', 'Child','Neutral']
for nametype  in NameType:
    itm["NameTypeCombo"].AddItem(nametype)

def on_language_combo_current_index_changed(ev):
    global lang
    itm["NameCombo"].Clear()
    lang_index = itm["LanguageCombo"].CurrentIndex

    selected_language = Language[lang_index]
    lang = next(locale for locale, data in voices.items() if data['language'] == selected_language)

    selected_type = itm["NameTypeCombo"].CurrentText

    matching_voices = [voice.split(' (')[0] for voice in voices[lang]['voices'] if voice.endswith(f"{selected_type})")]

    for voice_name in matching_voices:
        itm["NameCombo"].AddItem(voice_name)


def on_name_combo_current_index_changed(ev):
    itm["StyleCombo"].Clear()
    itm["StyleCombo"].AddItem('默认')
    itm["StyleCombo"].Enabled = False
    selected_voice = itm["NameCombo"].CurrentText.split(' (')[0]
    
    if selected_voice in styles:
        itm["StyleCombo"].Enabled = True
        for style in styles[selected_voice]:
            chinese_style=translate_styles_to_chinese(style)
            itm["StyleCombo"].AddItem(chinese_style)

    itm["MultilingualCombo"].Clear()
    itm["MultilingualCombo"].AddItem('默认')
    itm["MultilingualCombo"].Enabled = False
    if "Multilingual" in selected_voice:
        itm["MultilingualCombo"].Enabled = True
        for group_name, data in Multilinguals.items():
            cleaned_names = [n.split(' (')[0] for n in data["names"]]
            if selected_voice in cleaned_names:
                print(group_name)
                for language in data["languages"]:
                    chinese_lang = translate_lang_to_chinese(language)
                    itm["MultilingualCombo"].AddItem(chinese_lang)
                break 
      
   
    

def on_name_type_combo_current_index_changed(ev):
    itm["NameCombo"].Clear()
    selected_type = itm["NameTypeCombo"].CurrentText

    matching_voices = [voice.split(' (')[0] for voice in voices[lang]['voices'] if voice.endswith(f"{selected_type})")]

    for voice_name in matching_voices:
        itm["NameCombo"].AddItem(voice_name)


win.On.NameCombo.CurrentIndexChanged = on_name_combo_current_index_changed
win.On.LanguageCombo.CurrentIndexChanged = on_language_combo_current_index_changed
win.On.NameTypeCombo.CurrentIndexChanged = on_name_type_combo_current_index_changed
if saved_settings:
    itm["ApiKey"].Text = saved_settings.get("API_KEY", default_settings["API_KEY"])
    itm["Path"].Text = saved_settings.get("OUTPUT_DIRECTORY", default_settings["OUTPUT_DIRECTORY"])
    itm["Region"].Text = saved_settings.get("REGION", default_settings["REGION"])
    itm["LanguageCombo"].CurrentIndex = saved_settings.get("LANGUAGE", default_settings["LANGUAGE"])
    itm["NameTypeCombo"].CurrentIndex = saved_settings.get("TYPE", default_settings["TYPE"])
    itm["NameCombo"].CurrentIndex = saved_settings.get("NAME", default_settings["NAME"])
    itm["RateSpinBox"].Value = saved_settings.get("RATE", default_settings["RATE"])
    itm["PitchSpinBox"].Value = saved_settings.get("PITCH", default_settings["PITCH"])
    itm["VolumeSpinBox"].Value = saved_settings.get("VOLUME", default_settings["VOLUME"])
    itm["StyleCombo"].CurrentIndex = saved_settings.get("STYLE", default_settings["STYLE"])
    itm["StyleDegreeSpinBox"].Value = saved_settings.get("STYLEDEGREE", default_settings["STYLEDEGREE"])
    itm["OutputFormatCombo"].CurrentIndex = saved_settings.get("OUTPUT_FORMATS", default_settings["OUTPUT_FORMATS"])

def on_getsub_button_clicked(ev):
    subtitle = ''
    
    track_count = currentTimeline.GetTrackCount("subtitle")

    for track_index in range(1, track_count + 1):
        track_enabled = currentTimeline.GetIsTrackEnabled("subtitle", track_index)
        if track_enabled:
            subtitleTrackItems = currentTimeline.GetItemListInTrack("subtitle", track_index)
            
            for item in subtitleTrackItems:
                sub = item.GetName()
                subtitle += sub + "\n"
    
    itm["SubtitleTxt"].Text = subtitle
    print(subtitle)

win.On.GetSubButton.Clicked = on_getsub_button_clicked

    
def create_ssml(lang, voice_name, text, rate=None, pitch=None, volume=None, style=None, styledegree=None, multilingual='默认'):
    speak = ET.Element('speak', xmlns="http://www.w3.org/2001/10/synthesis",
                       attrib={
                           "xmlns:mstts": "http://www.w3.org/2001/mstts",
                           "xmlns:emo": "http://www.w3.org/2009/10/emotionml",
                           "version": "1.0",
                           "xml:lang": f"{lang}"
                       })

    voice = ET.SubElement(speak, 'voice', name=voice_name)
    
    if multilingual != "默认":
        lang_tag = ET.SubElement(voice, 'lang', attrib={"xml:lang": multilingual})
        parent_tag = lang_tag
    else:
        parent_tag = voice

    lines = text.split('\n')
    for line in lines:
        if line.strip():  # 确保不处理空行
            paragraph = ET.SubElement(parent_tag, 's')  # 创建段落元素
            if style!= "默认":
                express_as_attribs = {'style': style}
                if styledegree is not None and styledegree != 1.0:
                    express_as_attribs['styledegree'] = f"{styledegree:.2f}"
                express_as = ET.SubElement(paragraph, 'mstts:express-as', attrib=express_as_attribs)

                prosody_attrs = {}
                if rate is not None and rate != 1.0:
                    prosody_rate = f"+{(rate-1)*100:.2f}%" if rate > 1 else f"-{(1-rate)*100:.2f}%"
                    prosody_attrs['rate'] = prosody_rate
                if pitch is not None and pitch != 1.0:
                    prosody_pitch = f"+{(pitch-1)*100:.2f}%" if pitch > 1 else f"-{(1-pitch)*100:.2f}%"
                    prosody_attrs['pitch'] = prosody_pitch
                if volume is not None and volume != 1.0:
                    prosody_volume = f"+{(volume-1)*100:.2f}%" if volume > 1 else f"-{(1-volume)*100:.2f}%"
                    prosody_attrs['volume'] = prosody_volume

                if prosody_attrs:
                    prosody = ET.SubElement(express_as, 'prosody', attrib=prosody_attrs)
                    process_text_with_breaks(prosody, line)
                else:
                    process_text_with_breaks(express_as, line)
            else:
                prosody_attrs = {}
                if rate is not None and rate != 1.0:
                    prosody_rate = f"+{(rate-1)*100:.2f}%" if rate > 1 else f"-{(1-rate)*100:.2f}%"
                    prosody_attrs['rate'] = prosody_rate
                if pitch is not None and pitch != 1.0:
                    prosody_pitch = f"+{(pitch-1)*100:.2f}%" if pitch > 1 else f"-{(1-pitch)*100:.2f}%"
                    prosody_attrs['pitch'] = prosody_pitch
                if volume is not None and volume != 1.0:
                    prosody_volume = f"+{(volume-1)*100:.2f}%" if volume > 1 else f"-{(1-volume)*100:.2f}%"
                    prosody_attrs['volume'] = prosody_volume

                if prosody_attrs:
                    prosody = ET.SubElement(paragraph, 'prosody', attrib=prosody_attrs)
                    process_text_with_breaks(prosody, line)
                else:
                    process_text_with_breaks(paragraph, line)

    if multilingual:
        parent_tag.tail = "\n"
    
    return format_xml(ET.tostring(speak, encoding='unicode'))

def process_text_with_breaks(parent, text):
    parts = text.split('<break')
    for i, part in enumerate(parts):
        if i == 0:
            parent.text = part
        else:
            # 查找结束符
            end_idx = part.find('>')
            if end_idx != -1:
                break_tag = '<break' + part[:end_idx + 1]
                remaining_text = part[end_idx + 1:]
                
                # 插入 break 标签
                break_elem = ET.fromstring(break_tag)
                parent.append(break_elem)
                
                # 插入后续文本
                if remaining_text:
                    if len(parent) > 0 and parent[-1].tail is None:
                        parent[-1].tail = remaining_text
                    else:
                        parent.text = remaining_text

def format_xml(xml_string):
    parsed = minidom.parseString(xml_string)
    pretty_xml_as_string = parsed.toprettyxml(indent="  ")
    pretty_xml_as_string = '\n'.join([line for line in pretty_xml_as_string.split('\n') if line.strip()])
    return pretty_xml_as_string


def show_warning_message(text):
    msgbox = dispatcher.AddWindow(
        {
            "ID": 'msg',
            "WindowTitle": 'Warning',
            "Geometry": [750, 400, 350, 100],
            "Spacing": 10,
        },
        [
            ui.VGroup(
                [
                    ui.Label({"ID": 'WarningLabel', "Text": text}),
                    ui.HGroup(
                        {
                            "Weight": 0,
                        },
                        [
                            ui.Button({"ID": 'OkButton', "Text": 'OK'}),
                        ]
                    ),
                ]
            ),
        ]
    )

    def on_ok_button_clicked(ev):
        dispatcher.ExitLoop()
    msgbox.On.OkButton.Clicked = on_ok_button_clicked

    msgbox.Show()
    dispatcher.RunLoop()
    msgbox.Hide()
    
def update_status(message):
    itm["StatusLabel"].Text = message

def synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, volume, style, style_degree, multilingual,pitch,audio_format, audio_output_config):
    speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)
    speech_config.set_speech_synthesis_output_format(audio_format)
    
    ssml = create_ssml(lang=lang, voice_name=voice_name, text=subtitle, rate=rate, volume=volume, style=style, styledegree=style_degree,multilingual= multilingual,pitch=pitch)
    print(ssml)
    
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_output_config)
    result = speech_synthesizer.speak_ssml_async(ssml).get()
    
    return result

def on_play_button_clicked(ev):
    if itm["Path"].Text == '':
        show_warning_message('前往配置栏选择保存路径.')
        return
    if itm["ApiKey"].Text == '' or itm["Region"].Text == '':
        show_warning_message('前往配置栏填写区域和API密钥.')
        return
    
    update_status("播放中 ...")
    itm["PlayButton"].Text = "⏸"
    itm["PlayButton"].Enabled = False
    global subtitle, ssml, stream 
    service_region = itm["Region"].Text
    speech_key = itm["ApiKey"].Text
    style = get_original_style(itm["StyleCombo"].CurrentText)
    style_degree = itm["StyleDegreeSpinBox"].Value
    subtitle = itm["SubtitleTxt"].PlainText
    rate = itm["RateSpinBox"].Value
    pitch = itm["PitchSpinBox"].Value
    volume = itm["VolumeSpinBox"].Value
    multilingual = get_original_lang(itm["MultilingualCombo"].CurrentText)
    voice_name = itm["NameCombo"].CurrentText
    output_format = itm["OutputFormatCombo"].CurrentText
    if output_format in audio_formats:
        audio_format = audio_formats[output_format]
    else:
        show_warning_message('Unsupported audio format selected.')
        return

    audio_output_config = speechsdk.audio.AudioOutputConfig(use_default_speaker=True)
    
    result = synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, volume, style, style_degree, multilingual,pitch,audio_format, audio_output_config)
    stream = speechsdk.AudioDataStream(result)

    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        itm["PlayButton"].Text = "播放预览"
        update_status("")
    elif result.reason == speechsdk.ResultReason.Canceled:
        itm["PlayButton"].Text = "播放预览"
        cancellation_details = result.cancellation_details
        update_status(f"合成失败")
        print("Speech synthesis canceled: {}".format(cancellation_details.reason))
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print("Error details: {}".format(cancellation_details.error_details))
            
    itm["PlayButton"].Enabled = True  

win.On.PlayButton.Clicked = on_play_button_clicked

def on_load_button_clicked(ev):
    count = 0
    global current_context,subtitle,stream
    while True:
        count += 1
        filename = itm["Path"].Text + f"/Untitled#{count}" + itm["OutputFormatCombo"].CurrentText.split(", ")[1]
        if not os.path.exists(filename):
            break
    current_context = itm["SubtitleTxt"].PlainText
    print(f"stream:{stream}")
    if stream and current_context==subtitle:
        stream.save_to_wav_file(filename)
        time.sleep(1)
        add_to_media_pool(filename)
        update_status("加载成功")
        stream =None
    elif not stream and current_context!=subtitle:
        update_status("合成中 ...")
        service_region = itm["Region"].Text
        speech_key = itm["ApiKey"].Text
        style = get_original_style(itm["StyleCombo"].CurrentText)
        style_degree = itm["StyleDegreeSpinBox"].Value
        subtitle = itm["SubtitleTxt"].PlainText
        rate = itm["RateSpinBox"].Value
        pitch = itm["PitchSpinBox"].Value
        volume = itm["VolumeSpinBox"].Value
        multilingual = get_original_lang(itm["MultilingualCombo"].CurrentText)
        voice_name = itm["NameCombo"].CurrentText
        output_format = itm["OutputFormatCombo"].CurrentText
        if output_format in audio_formats:
            audio_format = audio_formats[output_format]
        else:
            show_warning_message('Unsupported audio format selected.')
            return

        audio_output_config = speechsdk.audio.AudioOutputConfig(filename=filename)
        
        result = synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, volume, style, style_degree, multilingual,pitch,audio_format, audio_output_config)
        
        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            time.sleep(1)
            add_to_media_pool(filename)
            update_status("加载成功")
        elif result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = result.cancellation_details
            update_status(f"合成失败")
            print("Speech synthesis canceled: {}".format(cancellation_details.reason))
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                print("Error details: {}".format(cancellation_details.error_details))
    else:
        update_status(f"媒体池已存在该片段")
        #current_context = itm["SubtitleTxt"].PlainText

win.On.LoadButton.Clicked = on_load_button_clicked

def on_break_button_clicked(ev):
    breaktime =  itm["BreakSpinBox"].Value
    # 插入<break>标志
    itm["SubtitleTxt"].InsertPlainText(f'<break time="{breaktime}ms" />')

win.On.BreakButton.Clicked = on_break_button_clicked



def on_reset_button_clicked(ev):
    #itm["ApiKey"].Text = default_settings["API_KEY"]
    #itm["Path"].Text = default_settings["OUTPUT_DIRECTORY"]
    #itm["Region"].Text = default_settings["REGION"]
    itm["LanguageCombo"].CurrentIndex = default_settings["LANGUAGE"]
    itm["NameTypeCombo"].CurrentIndex = default_settings["TYPE"]
    itm["NameCombo"].CurrentIndex = default_settings["NAME"]
    itm["RateSpinBox"].Value = default_settings["RATE"]
    itm["BreakSpinBox"].Value = default_settings["BREAKTIME"]
    itm["PitchSpinBox"].Value = default_settings["PITCH"]
    itm["VolumeSpinBox"].Value = default_settings["VOLUME"]
    itm["StyleCombo"].CurrentIndex = default_settings["STYLE"]
    itm["StyleDegreeSpinBox"].Value = default_settings["STYLEDEGREE"]
    itm["OutputFormatCombo"].CurrentIndex = default_settings["OUTPUT_FORMATS"]
win.On.ResetButton.Clicked = on_reset_button_clicked


def on_browse_button_clicked(ev):
    current_path = itm["Path"].Text
    selected_path = fusion.RequestDir(current_path)
    if selected_path:
        itm["Path"].Text = str(selected_path)
    else:
        print("No directory selected or the request failed.")
win.On.Browse.Clicked = on_browse_button_clicked

def close_and_save(settings_file):
    settings = {
        "API_KEY": itm["ApiKey"].Text,
        "REGION": itm["Region"].Text,
        "OUTPUT_DIRECTORY": itm["Path"].Text,
        "LANGUAGE": itm["LanguageCombo"].CurrentIndex,
        "TYPE": itm["NameTypeCombo"].CurrentIndex,
        "NAME": itm["NameCombo"].CurrentIndex,
        "RATE": itm["RateSpinBox"].Value,
        "PITCH": itm["PitchSpinBox"].Value,
        "VOLUME": itm["VolumeSpinBox"].Value,
        "STYLE": itm["StyleCombo"].CurrentIndex,
        "STYLEDEGREE": itm["StyleDegreeSpinBox"].Value,
        "OUTPUT_FORMATS": itm["OutputFormatCombo"].CurrentIndex,
    }

    save_settings(settings, settings_file)

def on_open_link_button_clicked(ev):
    webbrowser.open("https://mp.weixin.qq.com/s?__biz=MzUzMTk2MDU5Nw==&mid=2247484556&idx=1&sn=d4969f2d582c134d071aac4c23eb1df8&chksm=fabbc2f6cdcc4be085216b63b68cd772833e47bcdca90e1630e06a1957839729d5830e7321f8#rd")
win.On.OpenLinkButton.Clicked = on_open_link_button_clicked

def on_close(ev):
    close_and_save(settings_file)
    dispatcher.ExitLoop()
win.On.MyWin.Close = on_close


win.Show()
dispatcher.RunLoop()
win.Hide()

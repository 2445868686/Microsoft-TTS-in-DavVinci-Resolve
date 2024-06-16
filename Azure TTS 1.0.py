import platform
from xml.dom import minidom
import xml.etree.ElementTree as ET
import sys
import os
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
    "STYLE": 0,
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

    # 检查 TTS 文件夹是否已存在
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
    <h3>Introduction</h3>
    <p>This script uses Azure's TTS feature to convert text to speech.</p>

    <h3>Save Path</h3>
    <p>Specifies the save path for the generated files.</p>

    <h3>Region、API Key</h3>
    <p>Obtain your API key from <a href="https://speech.microsoft.com/">Microsoft Speech Studio</a></p>
</body>
</html>

"""

"""
<body>
    <h3>介绍</h3>
    <p>此脚本使用Azure的TTS功能将文字转换为语音。</p>

    <h3>保存路径</h3>
    <p>指定生成文件的保存路径。</p>

    <h3>区域、API密钥</h3>
    <p>从<a href="https://speech.microsoft.com/">Microsoft Speech Studio</a>获取您的API密钥。</p>
</body>
"""

ui = fusion.UIManager
dispatcher = bmd.UIDispatcher(ui)
win = dispatcher.AddWindow(
    {
        "ID": 'MyWin',
        "WindowTitle": 'TTS 1.0',
        "Geometry": [700, 300, 400, 480],
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
                                        ui.Button({"ID": 'GetSubButton', "Text": 'GetSubtitle'}),

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
                                  
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'StyleLabel', "Text": 'Style', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.ComboBox({"ID": 'StyleCombo', "Text": '', "Weight": 0.3}),
                                       ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'styledegreeLabel', "Text": 'StyleDegree', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'StyleDegreeSlider', "Value": 100, "Minimum": 0, "Maximum": 300,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'StyleDegreeSpinBox', "Value": 1.0, "Minimum": 0.0, "Maximum": 3.0, "SingleStep": 0.01, "Weight": 0.3}),
                                    ]
                                ),
                                ui.HGroup(
                                    {"Weight": 0.1},
                                    [
                                        
                                        ui.Label({"ID": 'rateLabel', "Text": 'Rate', "Alignment": {"AlignRight": False}, "Weight": 0.2}),
                                        ui.Slider({"ID": 'RateSlider', "Value": 100, "Minimum": 0, "Maximum": 300,  "Orientation": "Horizontal", "Weight": 0.5}),
                                        ui.DoubleSpinBox({"ID": 'RateSpinBox', "Value": 1.0, "Minimum": 0.0, "Maximum": 3.0, "SingleStep": 0.01, "Weight": 0.3}),
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

"""# 汉化界面
itm["GetSubButton"].Text = "从时间线获取字幕"
itm["LanguageLabel"].Text = "语言"
itm["NameTypeLabel"].Text = "类型"
itm["NameLabel"].Text = "名称"
itm["StyleLabel"].Text = "风格"
itm["styledegreeLabel"].Text = "风格强度"
itm["rateLabel"].Text = "速度"
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
"""
itm["MyTabs"].AddTab("Azure TTS")
itm["MyTabs"].AddTab("Configuration")

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
ssml = ''
voice_name = ""
style = None
rate = None
style_degree = None
stream = None

def on_subtitle_text_changed(ev):
    global stream
    stream = None
win.On.SubtitleTxt.TextChanged = on_subtitle_text_changed

Language = ['Chinese','English', 'German']
for language in Language:
    itm["LanguageCombo"].AddItem(language)

NameType = ['Female','Male', 'Child','Neutral']
for nametype  in NameType:
    itm["NameTypeCombo"].AddItem(nametype)


voices = {
    'zh-CN': [
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
    ],
    'en-US': [
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
    ],
    'de-DE': [
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

def on_language_combo_current_index_changed(ev):
    global lang
    itm["NameCombo"].Clear()
    lang_index = itm["LanguageCombo"].CurrentIndex

    if lang_index == 0:
        lang = 'zh-CN'
    elif lang_index == 1:
        lang = 'en-US'
    elif lang_index == 2:
        lang = 'de-DE'

    for voice in voices.get(lang, []):
        voice_name, voice_type = voice.split(' (')
        voice_type = voice_type.rstrip(')')
        itm["NameCombo"].AddItem(voice_name)

def on_name_combo_current_index_changed(ev):
    global voice_name
    itm["StyleCombo"].Clear()
    selected_voice = itm["NameCombo"].CurrentText.split(' (')[0]
    
    if selected_voice in styles:
        for style in styles[selected_voice]:
            itm["StyleCombo"].AddItem(style)
    voice_name = itm["NameCombo"].CurrentText

def on_name_type_combo_current_index_changed(ev):
    itm["NameCombo"].Clear()
    type_index = itm["NameTypeCombo"].CurrentIndex
    selected_type = itm["NameTypeCombo"].CurrentText

    for voice in voices.get(lang, []):
        voice_name, voice_type = voice.split(' (')
        voice_type = voice_type.rstrip(')')
        if voice_type == selected_type:
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
    itm["StyleCombo"].CurrentIndex = saved_settings.get("STYLE", default_settings["STYLE"])
    itm["StyleDegreeSpinBox"].Value = saved_settings.get("STYLEDEGREE", default_settings["STYLEDEGREE"])
    itm["OutputFormatCombo"].CurrentIndex = saved_settings.get("OUTPUT_FORMATS", default_settings["OUTPUT_FORMATS"])

def on_getsub_button_clicked(ev):
    subtitle = ''
    subtitleTrackIndex = 1  
    subtitleTrackItems = currentTimeline.GetItemListInTrack("subtitle", subtitleTrackIndex)

    for item in subtitleTrackItems:
        sub = item.GetName()
        subtitle += sub + "\n"

    itm["SubtitleTxt"].Text = subtitle
    print(subtitle)
win.On.GetSubButton.Clicked = on_getsub_button_clicked

    
def create_ssml(lang, voice_name, text, rate=None, style=None, styledegree=None):
    speak = ET.Element('speak', xmlns="http://www.w3.org/2001/10/synthesis",
                       attrib={
                           "xmlns:mstts": "http://www.w3.org/2001/mstts",
                           "xmlns:emo": "http://www.w3.org/2009/10/emotionml",
                           "version": "1.0",
                           "xml:lang": f"{lang}"
                       })

    voice = ET.SubElement(speak, 'voice', name=voice_name)
    
    lines = text.split('\n')
    for line in lines:
        if line.strip():  # 确保不处理空行
            paragraph = ET.SubElement(voice, 's')  # 创建段落元素
            if style:
                express_as_attribs = {'style': style}
                if styledegree is not None and styledegree != 1.0:
                    express_as_attribs['styledegree'] = f"{styledegree:.2f}"
                express_as = ET.SubElement(paragraph, 'mstts:express-as', attrib=express_as_attribs)

                if rate is not None and rate != 1.0:
                    prosody_rate = f"+{(rate-1)*100:.2f}%" if rate > 1 else f"-{(1-rate)*100:.2f}%"
                    prosody = ET.SubElement(express_as, 'prosody', rate=prosody_rate)
                    prosody.text = line
                else:
                    express_as.text = line
            else:
                if rate is not None and rate != 1.0:
                    prosody_rate = f"+{(rate-1)*100:.2f}%" if rate > 1 else f"-{(1-rate)*100:.2f}%"
                    prosody = ET.SubElement(paragraph, 'prosody', rate=prosody_rate)
                    prosody.text = line
                else:
                    paragraph.text = line

    return format_xml(ET.tostring(speak, encoding='unicode'))


def format_xml(xml_string):
    parsed = minidom.parseString(xml_string)
    pretty_xml_as_string = parsed.toprettyxml(indent="  ")
    pretty_xml_as_string = '\n'.join(pretty_xml_as_string.split('\n')[1:])
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

def synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, style, style_degree, audio_format, audio_output_config):
    speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)
    speech_config.set_speech_synthesis_output_format(audio_format)
    
    ssml = create_ssml(lang=lang, voice_name=voice_name, text=subtitle, rate=rate, style=style, styledegree=style_degree)
    print(ssml)
    
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_output_config)
    result = speech_synthesizer.speak_ssml_async(ssml).get()
    
    return result

def on_play_button_clicked(ev):
    if itm["Path"].Text == '':
        show_warning_message('Please go to Configuration to select the audio save path.')
        return
    if itm["ApiKey"].Text == '' or itm["Region"].Text == '':
        show_warning_message('Please go to Configuration to enter the API Key.')
        return
    
    update_status("Playing ...")
    #itm["PlayButton"].Text = "⏸"
    global subtitle, ssml, stream 
    service_region = itm["Region"].Text
    speech_key = itm["ApiKey"].Text
    style = itm["StyleCombo"].CurrentText
    style_degree = itm["StyleDegreeSpinBox"].Value
    subtitle = itm["SubtitleTxt"].PlainText
    rate = itm["RateSpinBox"].Value

    output_format = itm["OutputFormatCombo"].CurrentText
    if output_format in audio_formats:
        audio_format = audio_formats[output_format]
    else:
        show_warning_message('Unsupported audio format selected.')
        return

    audio_output_config = speechsdk.audio.AudioOutputConfig(use_default_speaker=True)
    
    result = synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, style, style_degree, audio_format, audio_output_config)
    stream = speechsdk.AudioDataStream(result)

    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        #itm["PlayButton"].Text = "Play"
        update_status("")
    elif result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = result.cancellation_details
        update_status(f"Failed to generate")
        print("Speech synthesis canceled: {}".format(cancellation_details.reason))
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print("Error details: {}".format(cancellation_details.error_details))

win.On.PlayButton.Clicked = on_play_button_clicked

def on_load_button_clicked(ev):
    count = 0
    while True:
        count += 1
        filename = itm["Path"].Text + f"/Untitled#{count}" + itm["OutputFormatCombo"].CurrentText.split(", ")[1]
        if not os.path.exists(filename):
            break

    if stream:
        stream.save_to_wav_file(filename)
        add_to_media_pool(filename)
        update_status("Load successful")
    else:
        update_status("Generate ...")
        service_region = itm["Region"].Text
        speech_key = itm["ApiKey"].Text
        style = itm["StyleCombo"].CurrentText
        style_degree = itm["StyleDegreeSpinBox"].Value
        subtitle = itm["SubtitleTxt"].PlainText
        rate = itm["RateSpinBox"].Value

        output_format = itm["OutputFormatCombo"].CurrentText
        if output_format in audio_formats:
            audio_format = audio_formats[output_format]
        else:
            show_warning_message('Unsupported audio format selected.')
            return

        audio_output_config = speechsdk.audio.AudioOutputConfig(filename=filename)
        
        result = synthesize_speech(service_region, speech_key, lang, voice_name, subtitle, rate, style, style_degree, audio_format, audio_output_config)
        
        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            add_to_media_pool(filename)
            update_status("Load successful")
        elif result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = result.cancellation_details
            update_status(f"Failed to generate")
            print("Speech synthesis canceled: {}".format(cancellation_details.reason))
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                print("Error details: {}".format(cancellation_details.error_details))

win.On.LoadButton.Clicked = on_load_button_clicked


def on_reset_button_clicked(ev):
    #itm["ApiKey"].Text = default_settings["API_KEY"]
    #itm["Path"].Text = default_settings["OUTPUT_DIRECTORY"]
    #itm["Region"].Text = default_settings["REGION"]
    itm["LanguageCombo"].CurrentIndex = default_settings["LANGUAGE"]
    itm["NameTypeCombo"].CurrentIndex = default_settings["TYPE"]
    itm["NameCombo"].CurrentIndex = default_settings["NAME"]
    itm["RateSpinBox"].Value = default_settings["RATE"]
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
        "STYLE": itm["StyleCombo"].CurrentIndex,
        "STYLEDEGREE": itm["StyleDegreeSpinBox"].Value,
        "OUTPUT_FORMATS": itm["OutputFormatCombo"].CurrentIndex,
    }

    save_settings(settings, settings_file)

def on_open_link_button_clicked(ev):
    webbrowser.open("https://www.paypal.me/heiba2wk")
win.On.OpenLinkButton.Clicked = on_open_link_button_clicked

def on_close(ev):
    close_and_save(settings_file)
    dispatcher.ExitLoop()
win.On.MyWin.Close = on_close


win.Show()
dispatcher.RunLoop()
win.Hide()

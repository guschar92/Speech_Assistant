#encoding=utf-8

from __future__ import print_function


# FIX: pkg_resources.get_distribution("google-cloud-speech").version
# E:\virtual\speech\Lib\site-packages\google\cloud\speech_v1\gapic\speech_client.py
def pkg_resources_get_distribution(dist):
    if dist == "google-cloud-speech":
        return pkg_resources.Distribution(version="1.0.0")
    return pkg_resources._get_distribution(dist)

import pkg_resources
pkg_resources._get_distribution = pkg_resources.get_distribution
pkg_resources.get_distribution = pkg_resources_get_distribution


# FIX: Exception  in 'grpc._cython.cygrpc.ssl_roots_override_callback' ignored
# E0411 15:40:32.864000000  9872 src/core/lib/security/security_connector/ssl_utils.cc:449] assertion failed: pem_root_certs != nullptr
import utils
import os
# E:\virtual\speech\Lib\site-packages\grpc\_cython\_credentials
os.environ["GRPC_DEFAULT_SSL_ROOTS_FILE_PATH"] = utils.resource_path("grpc.pem")


from google.cloud import speech
from google.cloud.speech import enums, types
from google.oauth2 import service_account
from google.api_core.exceptions import OutOfRange, ServiceUnavailable, ResourceExhausted, GoogleAPICallError
import struct
import math
import ctypes
import threading
import pyaudio
# import pyautogui
import pyperclip
import time
import datetime
import wx
import wx.adv
import os
import sys
import win32api
import win32con
import win32gui
import subprocess
import speech_recognition as sr
from microphone_stream import MicrophoneStream

# Audio recording parameters
RATE = 16000
CHUNK = int(RATE / 10)  # 100ms


class IntermediateFrame(wx.Frame):
    ALPHA = 180

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, style=wx.BORDER_NONE & (~wx.CLOSE_BOX) & (~wx.MAXIMIZE_BOX) & (~wx.RESIZE_BORDER))
        self.parent = parent
        self.cascadeClose = True
        self.ToggleWindowStyle(wx.STAY_ON_TOP)
        self.SetBackgroundColour(wx.Colour(0x000000))
        self.SetTransparent(self.ALPHA)

        app_icon_file = utils.resource_path("app_icon.ico")
        if os.path.isfile(app_icon_file):
            self.SetIcon(wx.Icon(app_icon_file))

        font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.messageLabel = wx.StaticText(self, style=wx.ALIGN_CENTRE_HORIZONTAL)
        self.messageLabel.SetFont(font)
        self.messageLabel.SetForegroundColour(wx.Colour(0xFFFFFF))

        dw, dh = wx.DisplaySize()
        self.SetSize(dw / 2, 150)

        w, h = self.GetSize()
        self.SetPosition(((dw - w) / 2, (dh - h) - 50))

        sizer = wx.BoxSizer(wx.VERTICAL)
        self.SetSizer(sizer)
        self.Layout()
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, event):
        if self.cascadeClose:
            self.parent.Close()
        self.Destroy()

    def SetMessageText(self, message):
        w, h = self.GetSize()
        self.messageLabel.SetLabelText(message)
        self.messageLabel.Wrap(w)

    def GetMessageText(self):
        return self.messageLabel.GetLabelText()

    def Display(self, show=True):
        if show:
            self.ShowWithEffect(wx.SHOW_EFFECT_EXPAND, 250)
        else:
            self.HideWithEffect(wx.SHOW_EFFECT_EXPAND, 250)

class SpeechAssistantFrame(wx.Frame):
    FORMAT = pyaudio.paInt16
    RATE = 44100
    CHANNELS = 1
    FRAMES_PER_BUFFER = 4096

    def __init__(self, *args, **kwargs):
        self.client = kwargs.pop("client")
        kwargs["style"] = wx.DEFAULT_FRAME_STYLE & (~wx.MAXIMIZE_BOX) & (~wx.RESIZE_BORDER)
        super(SpeechAssistantFrame, self).__init__(*args, **kwargs)
        self.SetSize([260, 210])
        self.ToggleWindowStyle(wx.STAY_ON_TOP)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.SetBackgroundColour(wx.Colour(255, 255, 255))

        # recording gif creation
        self.gif = wx.adv.AnimationCtrl(self, wx.ID_ANY, size=(-1, 42), pos=(20, 12))
        self.gif.LoadFile(utils.resource_path("rec.gif"))
        self.gif.Play()
        self.gif.Hide()

        # timer
        self.timelabel = wx.StaticText(self, pos=(self.GetSize()[0] - 60, 10), label="0:00")
        font = wx.Font(13, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        self.timelabel.SetFont(font)

        self.liveCheckBox = wx.CheckBox(self)
        self.liveCheckBox.SetLabelText(_("LIVE (real-time)"))
        font = self.liveCheckBox.GetFont()
        font = wx.Font(font.PixelSize, font.Family, font.Style, wx.BOLD)
        self.liveCheckBox.SetFont(font)
        if not self.client:
            #self.liveCheckBox.Disable()
            #self.liveCheckBox.SetToolTip("\"credentials.json\" not found")
            self.liveCheckBox.Hide()
            self.SetSize([260, 180])

        self.controlHoldingCheckBox = wx.CheckBox(self)
        self.controlHoldingCheckBox.SetLabelText(_("Hold right control to record"))

        self.middleClickCheckBox = wx.CheckBox(self)
        self.middleClickCheckBox.SetLabelText(_("Middle click to record"))
        #self.middleClickCheckBox.Bind(wx.EVT_CHECKBOX, self.OnMiddleClickModeChanged)

        self.upperCaseCheckBox = wx.CheckBox(self)
        self.upperCaseCheckBox.SetLabelText(_("Capital letters"))
        state_capital = win32api.GetKeyState(win32con.VK_CAPITAL)
        self.upperCaseCheckBox.SetValue(state_capital == 1)

        #self.moveCursorCheckBox = wx.CheckBox(self)
        #self.moveCursorCheckBox.SetLabelText(_("Move mouse to initial position"))
        #self.moveCursorCheckBox.SetValue(False)
        #self.moveCursorCheckBox.Disable()
        #self.moveCursorCheckBox.Hide()

        self.recordBtn = wx.Button(self)
        self.recordBtn.SetLabelText(_("Record"))
        self.recordBtn.Bind(wx.EVT_BUTTON, self.OnRecord)
        self.recordBtn.SetFocus()

        self.progress = wx.Gauge(self, range=100, style=wx.GA_HORIZONTAL, size=(-1, 4))
        self.label = wx.StaticText(self, -1, label=_("Start recording\nby pressing right Control!"), style=wx.ALIGN_CENTER | wx.ST_NO_AUTORESIZE)
        # self.label.Wrap(200)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.recordBtn, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.progress, 0, wx.EXPAND)
        sizer.Add(self.label, 0, wx.EXPAND, 5)
        sizer.Add(self.liveCheckBox, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.upperCaseCheckBox, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.controlHoldingCheckBox, 0, wx.ALL | wx.CENTER, 5)
        sizer.Add(self.middleClickCheckBox, 0, wx.ALL | wx.CENTER, 5)
        #sizer.Add(self.moveCursorCheckBox, 0, wx.ALL | wx.CENTER, 5)
        self.SetSizer(sizer)

        self.lastClick = None
        self.lastTime = None
        self.intermediateMessage = None
        self.recording = False
        self.audioSource = pyaudio.PyAudio()
        self.frames = []
        self.audioList = []
        self.lock = threading.Lock()
        self.threads = []
        self.stream = None
        dw, dh = wx.DisplaySize()
        w, h = self.GetSize()
        x = dw - w
        y = dh - h - 40
        self.SetPosition((x, y))

        self.log(self.getDate() + " - Application Started\n")

        self.recognizer = sr.Recognizer()
        self.recognizer.phrase_threshold = 0.3
        self.recognizer.pause_threshold = 0.2
        self.recognizer.non_speaking_duration = 0.1

        self.intermediateFrame = IntermediateFrame(self)
        #self.intermediateFrame.Show()

        self.prefix = ""
        self.stopper = None
        self.recordingStartTime = None
        self.loop = True
        th = threading.Thread(target=self.eventListener)
        th.start()

    #def OnMiddleClickModeChanged(self, event):
    #    if event.IsChecked():
    #       self.moveCursorCheckBox.Enable()
    #    else:
    #        self.moveCursorCheckBox.Disable()

    def log(self, msg):
        with open("speech_assistant_log.txt", "a") as f:
            f.write(msg)

    def getDate(self):
        return "[" + str(datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S")) + "]"

    def eventListener(self):
        state_capital = win32api.GetKeyState(win32con.VK_CAPITAL)
        state_control = win32api.GetKeyState(win32con.VK_RCONTROL)
        state_middle = win32api.GetKeyState(win32con.VK_MBUTTON)  # button up = 0 or 1. Button down = -127 or -128
        while self.loop:
            if state_capital != win32api.GetKeyState(win32con.VK_CAPITAL):
                state_capital = win32api.GetKeyState(win32con.VK_CAPITAL)
                if state_capital >= 0:
                    self.upperCaseCheckBox.SetValue(state_capital == 1)

            if self.middleClickCheckBox.IsChecked():
                a = win32api.GetKeyState(win32con.VK_MBUTTON)
                if a != state_middle:  # Button state changed
                    state_middle = a
                    if a < 0:
                        pos = win32gui.GetCursorPos()
                        performClick(pos[0], pos[1])
                        if self.recording:
                            self.lastClick = pos
                        if self.recordBtn.IsEnabled():
                            self.OnRecord()

            a = win32api.GetKeyState(win32con.VK_RCONTROL)
            if a != state_control:  # Button state changed
                state_control = a
                if not self.controlHoldingCheckBox.IsChecked():
                    if a < 0 and self.recordBtn.IsEnabled():
                        self.OnRecord()
                    continue
                if a < 0 and not self.recording:
                    time.sleep(0.1)
                    a = win32api.GetKeyState(win32con.VK_RCONTROL)
                if (a >= 0 and self.recording) or (a < 0 and not self.recording):
                    pos = win32gui.GetCursorPos()
                    if self.recording:
                        self.lastClick = pos
                    if self.recordBtn.IsEnabled():
                        self.OnRecord()
            time.sleep(0.1)

    def OnRecord(self, event=None, rec=None):
        if rec is None:
            rec = self.recording
        if rec:
            self.gif.Hide()
            self.liveCheckBox.Enable()
            self.log(self.getDate() + " - Recording Stopped...\n")
            self.recording = False
            self.recordBtn.SetLabelText(_("Record"))
            self.label.SetLabelText(_("Start recording\nby pressing right Control!"))
            if self.stopper:
                self.stopper(False)
                self.stopper = None
            #for th in self.threads:
            #    th.join()
        else:
            self.gif.Show()
            self.liveCheckBox.Disable()
            self.recordingStartTime = time.time()
            self.log(self.getDate() + " - Recording Started...\n")
            self.recording = True
            self.recordBtn.SetLabelText(_("Stop"))
            self.label.SetLabelText(_("Recording...\nStop by pressing right control"))
            self.frames = []
            self.audioList = []
            self.prefix = ""
            self.threads = []
            if self.liveCheckBox.IsChecked():
                th = threading.Thread(target=self.live_recognize_loop)
                self.stopper = None
            else:
                th = threading.Thread(target=self.recognize_loop)
                self.stopper = None
                #self.stopper = self.recognizer.listen_in_background(sr.Microphone(), self.OnPhraseDetected)
            th.start()
            self.threads.append(th)

            th = threading.Thread(target=self.record_loop)
            th.start()
            self.threads.append(th)

    def OnPhraseDetected(self, recognizer, audio):
        with self.lock:
            self.audioList.append(audio)

    def recognize_loop(self):
        def is_running():
            return self.recording

        frame_data = None
        while self.recording:
            with MicrophoneStream(RATE, CHUNK) as stream:
                audio_generator = stream.generator(is_running)
                frame_data = b''.join(list(audio_generator))

        if frame_data:
            audio = sr.AudioData(frame_data, RATE, self.audioSource.get_sample_size(self.FORMAT))
            self.recordBtn.Disable()

            self.log(self.getDate() + " - Recognizing...\n")
            self.timelabel.SetForegroundColour(wx.Colour(0xFFB000))
            self.timelabel.SetLabelText(self.timelabel.GetLabelText())  # Redraw

            try:
                message = self.recognizer.recognize_google(audio, language="el-GR")
            except sr.UnknownValueError:
                message = None
                self.log(self.getDate() + " - NOT RECOGNIZED\n")
            except sr.RequestError as e:
                message = None
                self.log("{0} - NOT RECOGNIZED - {1}\n".format(self.getDate(), e))

            if message:
                self.log(self.getDate() + " - Recognized: '%s'\n" % message.encode('utf-8'))
                self.user_display(message)
            self.recordBtn.Enable()

        self.timelabel.SetForegroundColour(wx.Colour(0x000000))
        self.timelabel.SetLabelText(self.timelabel.GetLabelText())  # Redraw

    def recognize_loop_by_phrase(self):
        while self.recording:
            audio = None
            with self.lock:
                if self.audioList:
                    audio = self.audioList.pop()
            if audio:
                self.log(self.getDate() + " - Recognizing...\n")
                self.timelabel.SetForegroundColour(wx.Colour(0xFFB000))
                try:
                    message = self.recognizer.recognize_google(audio, language="el-GR")
                except sr.UnknownValueError:
                    message = None
                    self.log(self.getDate() + " - NOT RECOGNIZED\n")
                except sr.RequestError as e:
                    message = None
                    self.log("{0} - NOT RECOGNIZED - {1}\n".format(self.getDate(), e))

                self.timelabel.SetForegroundColour(wx.Colour(0x000000))
                if message:
                    self.log(self.getDate() + " - Recognized: '%s'\n" % message.encode('utf-8'))
                    self.user_display(message)
            else:
                time.sleep(0.5)
        self.timelabel.SetForegroundColour(wx.Colour(0x000000))
        self.timelabel.SetLabelText(self.timelabel.GetLabelText())  # Redraw

    def live_recognize_loop(self):
        self.timelabel.SetForegroundColour(wx.Colour(0xFFB000))
        client = self.client

        def is_running():
            return self.recording

        while self.recording:
            #print("Start over")
            with MicrophoneStream(RATE, CHUNK) as stream:
                audio_generator = stream.generator(is_running)
                requests = (types.StreamingRecognizeRequest(audio_content=content) for content in audio_generator)
                responses = client.streaming_recognize(client.custom_streaming_config, requests)

                responses_iterator = iter(responses)
                while self.recording:
                    try:
                        response = next(responses_iterator)
                    except StopIteration:
                        break
                    except OutOfRange:
                        # Exception 400 - Exceeded maximum allowed stream duration of 65 seconds.
                        self.user_display(self.intermediateFrame.GetMessageText())
                        break  # Start over
                    except ServiceUnavailable as e:
                        # Exception 503 - Getting metadata from plugin failed
                        self.log("{0} - NOT RECOGNIZED - {1}\n".format(self.getDate(), e))
                        break
                    except ResourceExhausted as e:
                        # Exception 429 - Quota exceeded for quota metric 'speech.googleapis.com/default_requests' and
                        # limit 'DefaultRequestsPerMinutePerProject' of service 'speech.googleapis.com' for consumer
                        self.log("{0} - NOT RECOGNIZED - {1}\n".format(self.getDate(), e))
                        break
                    except GoogleAPICallError as e:
                        self.log("{0} - NOT RECOGNIZED - {1}\n".format(self.getDate(), e))
                        break

                    if response.results:
                        result = response.results[0]
                        if result.alternatives:
                            transcript = result.alternatives[0].transcript
                            self.intermediateFrame.SetMessageText(transcript)
                            if not result.is_final:
                                self.intermediateFrame.Display()
                                # print(transcript)
                            else:
                                self.user_display(transcript)
                                self.intermediateFrame.Display(False)
                                self.intermediateFrame.SetMessageText("")
                                #print("\t\t FINAL: %s" % transcript)
                                break  # Start over

        self.user_display(self.intermediateFrame.GetMessageText())
        self.intermediateFrame.Display(False)
        self.intermediateFrame.SetMessageText("")
        self.timelabel.SetForegroundColour(wx.Colour(0x000000))
        self.timelabel.SetLabelText(self.timelabel.GetLabelText())  # Redraw

    def user_display(self, message):
        if not message:
            return
        message = self.validate(message)
        message = self.toUpper(message)

        clipboard = pyperclip.paste()
        message = self.prefix + message
        pyperclip.copy(message)
        if not self.prefix:
            self.prefix = " "
        # pyautogui.hotkey("ctrl", "v")
        self.pressHoldRelease(win32con.VK_LCONTROL, ord("V"))  # ctrl + V
        pyperclip.copy(clipboard)

    def OnClose(self, event):
        self.log(self.getDate() + " - Application Terminated\n")
        self.recording = False
        self.progress = None
        self.loop = False
        time.sleep(0.2)
        self.intermediateFrame.cascadeClose = False
        self.intermediateFrame.Close()
        self.Destroy()

    def record_loop(self):
        # Possible interference with streaming recognition
        #audioSource = pyaudio.PyAudio()
        #stream = audioSource.open(format=self.FORMAT, channels=self.CHANNELS, rate=self.RATE, input=True, frames_per_buffer=self.FRAMES_PER_BUFFER)
        while self.recording:
            #data = stream.read(self.FRAMES_PER_BUFFER)
            #if self.progress:
            #    self.progress.SetValue(self.decibel(data))
            self.progress.SetValue(50)

            seconds = int(time.time() - self.recordingStartTime)
            minutes = int(seconds / 60)
            seconds = seconds % 60
            self.timelabel.SetLabel("%d:%02d" % (minutes, seconds))
            time.sleep(0.5)
        if self.progress:
            self.progress.SetValue(0)

    def validate(self, message):
        # All digits or white spaces
        if sum(c.isdigit() or c.isspace() or c == "." for c in message) == len(message):
            return message.replace(" ", "").replace(".", ",")
        # ID numbers
        if (sum(c.isalpha() for c in message) in [1, 2]) and sum(c.isdigit() for c in message) == 6:
            return message.replace(" ", "")
        # KAEK
        if sum(c.isdigit() for c in message) == 12:
            return message.replace(" ", "")
        # telephone numbers
        if sum(c.isdigit() for c in message) == 10:
            return message.replace(" ", "")
        # elif (u'παπάκι'.encode('utf-8') in message.encode('utf-8')):
        #     return message.replace(u'παπάκι ',"@")
        return message

    def toUpper(self, message):
        # checkbox to convert to uppercase
        if self.upperCaseCheckBox.IsChecked():
            message = message.upper()
            for ch1, ch2 in {"ά": "α", "έ": "ε", "ί": "ι", "ΐ": "ϊ", "ή": "η",
                             "ό": "ο", "ύ": "υ", "ΰ": "ϋ", "ώ": "ω"}.iteritems():
                message = message.replace(ch1.decode("utf-8").upper(), ch2.decode("utf-8").upper())
        return message

    def pressHoldRelease(self, *args):
        for i in args:
            win32api.keybd_event(i, 0, 0, 0)
            time.sleep(.05)

        for i in args:
            win32api.keybd_event(i, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)

    @staticmethod
    def decibel(data):
        count = len(data) / 2
        if count == 0:
            return 0
        sum_squares = 0
        shorts = struct.unpack("%dh" % count, data)
        for sample in shorts:
            n = sample * (1.0 / 32768)
            sum_squares += n * n
        rms = math.sqrt(sum_squares / count)
        return 100 + 20 * math.log10(rms)


def performClick(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


class ExtendedPopen(subprocess.Popen):
    def __init__(self, *args, **kwargs):
        for stream_name in ["stdin", "stdout", "stderr"]:
            if kwargs.get(stream_name) is None:
                kwargs[stream_name] = subprocess.PIPE
        super(ExtendedPopen, self).__init__(*args, **kwargs)


def _(msg):
    return {
        "Speech Assistant": "Αναγνώριση ομιλίας",
        "Middle click to record": "Ηχογράφηση πατώντας την ροδέλα",
        "Capital letters": "Κεφαλαία γράμματα",
        "Hold right control to record": "Ηχογράφηση κρατώντας το δεξί Control",
        "Move mouse to initial position": "Μετακίνηση ποντικιού στην αρχική θέση",
        "Record": "Ηχογράφηση",
        "Stop": "Διακοπή",
        "Start recording\nby pressing right Control!": "Ξεκινήστε την ηχογράφηση\nπατώντας το δεξί Control!",
        "Recording...\nStop by pressing right control": "Ηχογράφηση...\nΔιακοπή πατώντας το δεξί Control",
        "Stopped.": "Διακοπή.",
        "Stopped.\nRecognizing...": "Διακοπή.\nΓίνεται αναγνώριση...",
        "Stopped. %s\nPress right control": "%s\nΠατήστε το δεξί Control",
        "Recognition completed.": "Επιτυχής αναγνώριση.",
        "No data.": "Προσπαθήστε ξανά...",
        "LIVE (real-time)": "LIVE (σε πραγματικό χρόνο)"
    }.get(msg, msg)


def main():
    # Ensure that all streams are redirected
    subprocess.Popen = ExtendedPopen

    if len(sys.argv) > 0:
        script_name = sys.argv[0]
        if script_name.endswith(".exe"):
            # Stop other instances
            script_name = os.path.basename(script_name)
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            try:
                subprocess.check_call(('TASKKILL /F /IM ' + script_name + ' /FI').split(" ") + ["PID ne %s" % os.getpid()], startupinfo=startupinfo)
            except Exception:
                pass

    set_model_id = getattr(ctypes.windll.shell32, "SetCurrentProcessExplicitAppUserModelID", None)
    if callable(set_model_id):
        set_model_id('SoundRecordApplication')

    app_icon_file = utils.resource_path("app_icon.ico")
    credentials_file = utils.resource_path("credentials.json")
    if os.path.isfile(credentials_file):
        language_code = 'el-GR'  # a BCP-47 language tag
        credentials = service_account.Credentials.from_service_account_file(credentials_file)
        client = speech.SpeechClient(credentials=credentials)
        config = types.RecognitionConfig(
            encoding=enums.RecognitionConfig.AudioEncoding.LINEAR16,
            sample_rate_hertz=RATE,
            language_code=language_code)
        streaming_config = types.StreamingRecognitionConfig(
            config=config,
            interim_results=True,
            single_utterance=False)
        client.custom_streaming_config = streaming_config
    else:
        client = None

    app = wx.App()
    frm = SpeechAssistantFrame(None, title=_("Speech Assistant"), client=client)
    if os.path.isfile(app_icon_file):
        frm.SetIcon(wx.Icon(app_icon_file))

    frm.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()

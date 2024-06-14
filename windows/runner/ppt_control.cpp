#include "flutter_window.h"

#include "flutter/generated_plugin_registrant.h"

#include <string>

#include <flutter/binary_messenger.h>

#include <flutter/standard_method_codec.h>

#include <flutter/method_channel.h>

#include <flutter/method_result_functions.h>

#include <flutter/encodable_value.h>

#include <windows.h>

#include <atlbase.h>

#include <atlcom.h>

#include <comdef.h>

#include <atlcomcli.h>

#include <mshtml.h>

#include <mshtmdid.h>

#include <exdispid.h>

#include <exdisp.h>

#include <shellapi.h>

#include <iostream>


namespace ppt_control {
  class createChannelOpenApp {
    public: createChannelOpenApp(flutter::FlutterEngine * engine) {
      initialize(engine);
    }

    void initialize(flutter::FlutterEngine * FlEngine) {
      const static std::string channel_name("open_app_channel");
      flutter::BinaryMessenger * messenger = FlEngine -> messenger();
      const flutter::StandardMethodCodec * codec = & flutter::StandardMethodCodec::GetInstance();
      auto channel = std::make_unique < flutter::MethodChannel < >> (messenger, channel_name, codec);
      channel -> SetMethodCallHandler(
        [ & ](const flutter::MethodCall < > & call, std::unique_ptr < flutter::MethodResult < >> result) {
          AddMethodHandlers(call, & result);
        });
    }

    void AddMethodHandlers(const flutter::MethodCall < > & call, std::unique_ptr < flutter::MethodResult < >> * result) {
      if (call.method_name().compare("openPowerPoint") == 0) {
        try {
          handleOpenPowerPoint(call, result);
        } catch (...) {
          ( * result) -> Error("An error occurred");
        }
      } else if (call.method_name().compare("nextSlide") == 0) {
        try {
          handleNextSlide(call, result);
        } catch (...) {
          ( * result) -> Error("An error occurred");
        }
      } else if (call.method_name().compare("previousSlide") == 0) {
        try {
          handlePreviousSlide(call, result);
        } catch (...) {
          ( * result) -> Error("An error occurred");
        }
      } else {
        ( * result) -> NotImplemented();
      }
    }
    void handleOpenPowerPoint(const flutter::MethodCall < > & call, std::unique_ptr < flutter::MethodResult < >> * resPointer) {
      std::wstring powerPointPath = L"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE";
      HINSTANCE hInst = ShellExecuteW(NULL, L"open", powerPointPath.c_str(), NULL, NULL, SW_SHOWNORMAL);
      if ((uintptr_t) hInst > 32) {
        ( * resPointer) -> Success(flutter::EncodableValue("PowerPoint opened successfully"));
      } else {
        ( * resPointer) -> Error("Error", "Could not open PowerPoint");
      }
    }
    void handleNextSlide(const flutter::MethodCall < > & call, std::unique_ptr < flutter::MethodResult < >> * resPointer) {
      HWND hwnd = FindWindow(L"screenClass", NULL); // The class name for PowerPoint slideshow window

      if (hwnd != NULL) {
        // Bring the PowerPoint window to the foreground
        SetForegroundWindow(hwnd);
        SetActiveWindow(hwnd);

        // Simulate a "Right Arrow" key press and release event
        INPUT ip;
        ip.type = INPUT_KEYBOARD;
        ip.ki.wScan = 0;
        ip.ki.time = 0;
        ip.ki.dwExtraInfo = 0;

        // Press the "Right Arrow" key
        ip.ki.wVk = VK_RIGHT; // virtual-key code for the "Right Arrow" key
        ip.ki.dwFlags = 0; // 0 for key press
        SendInput(1, & ip, sizeof(INPUT));

        // Release the "Right Arrow" key
        ip.ki.dwFlags = KEYEVENTF_KEYUP; // KEYEVENTF_KEYUP for key release
        SendInput(1, & ip, sizeof(INPUT));

        ( * resPointer) -> Success(flutter::EncodableValue("Next slide command sent"));
      } else {
        ( * resPointer) -> Error("Error", "PowerPoint slideshow window not found");
      }
    }
    void handlePreviousSlide(const flutter::MethodCall < > & call, std::unique_ptr < flutter::MethodResult < >> * resPointer) {
      HWND hwnd = FindWindow(L"screenClass", NULL); // The class name for PowerPoint slideshow window

      if (hwnd != NULL) {
        // Bring the PowerPoint window to the foreground
        SetForegroundWindow(hwnd);
        SetActiveWindow(hwnd);

        // Simulate a "Left Arrow" key press and release event
        INPUT ip;
        ip.type = INPUT_KEYBOARD;
        ip.ki.wScan = 0;
        ip.ki.time = 0;
        ip.ki.dwExtraInfo = 0;

        // Press the "Left Arrow" key
        ip.ki.wVk = VK_LEFT; // virtual-key code for the "Left Arrow" key
        ip.ki.dwFlags = 0; // 0 for key press
        SendInput(1, & ip, sizeof(INPUT));

        // Release the "Left Arrow" key
        ip.ki.dwFlags = KEYEVENTF_KEYUP; // KEYEVENTF_KEYUP for key release
        SendInput(1, & ip, sizeof(INPUT));

        ( * resPointer) -> Success(flutter::EncodableValue("Previous slide command sent"));
      } else {
        ( * resPointer) -> Error("Error", "PowerPoint slideshow window not found");
      }
    }
  };
}
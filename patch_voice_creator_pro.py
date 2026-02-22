#!/usr/bin/env python3
"""
Voice Creator Pro - Comprehensive Patcher
==========================================
Applies all fixes in a single run:
  1. Enable Windows Long Paths (registry)
  2. RTX 5080 / Blackwell GPU Support (Triton patches)
  3. Speed & Enthusiasm Sliders (UI + backend)
  4. Expanded Speaker List with timing fix

Run from an elevated (Administrator) command prompt:
    python patch_voice_creator_pro.py

A reboot is recommended after patching (required for the long paths registry change).
"""

import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import urllib.request
import zipfile


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROGRAM_FILES = os.environ.get("ProgramFiles", r"C:\Program Files")
LOCAL_APPDATA = os.environ.get("LOCALAPPDATA", os.path.expandvars(r"%LOCALAPPDATA%"))

VCP_DIR = os.path.join(PROGRAM_FILES, "Voice Creator Pro")
INTERNAL = os.path.join(VCP_DIR, "_internal")
UI_DIR = os.path.join(INTERNAL, "ui")
ASSETS_DIR = os.path.join(UI_DIR, "assets")
UTILS_PY = os.path.join(INTERNAL, "transformers", "generation", "utils.py")
INDEX_HTML = os.path.join(UI_DIR, "index.html")

PACKAGES_DIR = os.path.join(LOCAL_APPDATA, "VoiceCloner", "packages")
PYTHON_DEV_DIR = os.path.join(PACKAGES_DIR, "python_dev")
TRITON_BUILD_PY = os.path.join(PACKAGES_DIR, "triton", "runtime", "build.py")
TRITON_WINUTILS = os.path.join(PACKAGES_DIR, "triton", "windows_utils.py")

NUGET_URL = (
    "https://www.nuget.org/api/v2/package/python/3.11.9"
)

CSS_FILE = os.path.join(ASSETS_DIR, "voice-controls.css")
JS_FILE = os.path.join(ASSETS_DIR, "voice-controls.js")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def check_admin():
    """Return True if running with Administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def backup(path):
    """Create a .bak copy of a file (only if .bak doesn't already exist)."""
    bak = path + ".bak"
    if not os.path.exists(bak) and os.path.exists(path):
        shutil.copy2(path, bak)


def status(label, ok, detail=""):
    tag = "OK" if ok else "FAIL"
    msg = f"  [{tag}] {label}"
    if detail:
        msg += f"  ({detail})"
    print(msg)


# ---------------------------------------------------------------------------
# Fix 1: Enable Windows Long Paths
# ---------------------------------------------------------------------------
def fix_long_paths():
    """Set LongPathsEnabled = 1 in the registry."""
    print("\n[Fix 1] Enable Windows Long Paths")

    # Check current value
    try:
        result = subprocess.run(
            ["reg", "query",
             r"HKLM\SYSTEM\CurrentControlSet\Control\FileSystem",
             "/v", "LongPathsEnabled"],
            capture_output=True, text=True
        )
        if "0x1" in result.stdout:
            status("LongPathsEnabled", True, "already enabled")
            return True
    except Exception:
        pass

    try:
        subprocess.run(
            ["reg", "add",
             r"HKLM\SYSTEM\CurrentControlSet\Control\FileSystem",
             "/v", "LongPathsEnabled", "/t", "REG_DWORD", "/d", "1", "/f"],
            check=True, capture_output=True, text=True
        )
        status("LongPathsEnabled", True, "set to 1 -- reboot required")
        return True
    except subprocess.CalledProcessError as exc:
        status("LongPathsEnabled", False, str(exc))
        return False


# ---------------------------------------------------------------------------
# Fix 2: RTX 5080 / Blackwell GPU Support
# ---------------------------------------------------------------------------
def fix_gpu_triton():
    """Download Python 3.11.9 dev files and patch Triton's build system."""
    print("\n[Fix 2] RTX 5080 / Blackwell GPU Support (Triton patches)")
    all_ok = True

    # --- 2a: Download Python dev files from NuGet ---
    include_dir = os.path.join(PYTHON_DEV_DIR, "include")
    libs_dir = os.path.join(PYTHON_DEV_DIR, "libs")
    python_h = os.path.join(include_dir, "Python.h")
    python_lib = os.path.join(libs_dir, "python311.lib")

    if os.path.isfile(python_h) and os.path.isfile(python_lib):
        status("Python dev files", True, "already present")
    else:
        try:
            os.makedirs(PYTHON_DEV_DIR, exist_ok=True)
            nupkg = os.path.join(tempfile.gettempdir(), "python-3.11.9.nupkg")
            print("    Downloading Python 3.11.9 NuGet package...")
            urllib.request.urlretrieve(NUGET_URL, nupkg)

            with zipfile.ZipFile(nupkg, "r") as zf:
                # Extract include/ files
                os.makedirs(include_dir, exist_ok=True)
                for entry in zf.namelist():
                    if entry.startswith("tools/include/") and not entry.endswith("/"):
                        fname = os.path.basename(entry)
                        # Preserve subdirectory structure under include/
                        rel = entry[len("tools/include/"):]
                        dest = os.path.join(include_dir, rel)
                        os.makedirs(os.path.dirname(dest), exist_ok=True)
                        with zf.open(entry) as src, open(dest, "wb") as dst:
                            dst.write(src.read())

                # Extract libs/ files
                os.makedirs(libs_dir, exist_ok=True)
                for entry in zf.namelist():
                    if entry.startswith("tools/libs/") and not entry.endswith("/"):
                        fname = os.path.basename(entry)
                        dest = os.path.join(libs_dir, fname)
                        with zf.open(entry) as src, open(dest, "wb") as dst:
                            dst.write(src.read())

            os.remove(nupkg)

            if os.path.isfile(python_h) and os.path.isfile(python_lib):
                status("Python dev files", True, "downloaded from NuGet")
            else:
                status("Python dev files", False, "extraction incomplete")
                all_ok = False
        except Exception as exc:
            status("Python dev files", False, str(exc))
            all_ok = False

    # --- 2b: Patch triton/runtime/build.py ---
    if not os.path.isfile(TRITON_BUILD_PY):
        status("Triton build.py", False, "file not found")
        all_ok = False
    else:
        with open(TRITON_BUILD_PY, "r", encoding="utf-8") as f:
            content = f.read()

        PATCH_MARKER = "# Fallback for PyInstaller"
        if PATCH_MARKER in content:
            status("Triton build.py", True, "already patched")
        else:
            # Anchor: insert after the py_include_dir = sysconfig... line
            anchor = 'py_include_dir = sysconfig.get_paths(scheme=scheme)["include"]'
            if anchor not in content:
                status("Triton build.py", False, "anchor line not found")
                all_ok = False
            else:
                patch = '''\
    py_include_dir = sysconfig.get_paths(scheme=scheme)["include"]
    # Fallback for PyInstaller: if Python.h is missing, use bundled python_dev headers
    if not os.path.isfile(os.path.join(py_include_dir, "Python.h")):
        # Navigate from build.py's location: runtime/ -> triton/ -> packages/ -> python_dev/include/
        _packages_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        _fallback = os.path.join(_packages_dir, "python_dev", "include")
        if os.path.isfile(os.path.join(_fallback, "Python.h")):
            py_include_dir = _fallback'''
                backup(TRITON_BUILD_PY)
                content = content.replace(anchor, patch)
                with open(TRITON_BUILD_PY, "w", encoding="utf-8") as f:
                    f.write(content)
                status("Triton build.py", True, "patched")

    # --- 2c: Patch triton/windows_utils.py ---
    if not os.path.isfile(TRITON_WINUTILS):
        status("Triton windows_utils.py", False, "file not found")
        all_ok = False
    else:
        with open(TRITON_WINUTILS, "r", encoding="utf-8") as f:
            content = f.read()

        PATCH_MARKER = "# Fallback for PyInstaller"
        if PATCH_MARKER in content:
            status("Triton windows_utils.py", True, "already patched")
        else:
            # Anchor: the return [str(python_lib_dir)] line followed by warnings
            anchor = '            return [str(python_lib_dir)]\n\n    warnings.warn("Failed to find Python libs.")'
            if anchor not in content:
                status("Triton windows_utils.py", False, "anchor not found")
                all_ok = False
            else:
                patch = '''\
            return [str(python_lib_dir)]

    # Fallback for PyInstaller: check bundled python_dev libs
    # Navigate from windows_utils.py's location: triton/ -> packages/ -> python_dev/libs/
    _packages_dir = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    _fallback_lib_dir = _packages_dir / "python_dev" / "libs"
    if (_fallback_lib_dir / f"python{version}.lib").exists():
        return [str(_fallback_lib_dir)]

    warnings.warn("Failed to find Python libs.")'''
                backup(TRITON_WINUTILS)
                content = content.replace(anchor, patch)
                with open(TRITON_WINUTILS, "w", encoding="utf-8") as f:
                    f.write(content)
                status("Triton windows_utils.py", True, "patched")

    return all_ok


# ---------------------------------------------------------------------------
# Fix 3 + 4: Speed/Enthusiasm Sliders + Expanded Speaker List
# ---------------------------------------------------------------------------

VOICE_CONTROLS_CSS = """\
/* Voice Creator Pro - Speed & Enthusiasm Controls */
#vc-voice-controls {
  padding: 8px 0 4px 0;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.vc-slider-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.vc-slider-label {
  font-size: 11px;
  font-weight: 500;
  color: #71717a;
  min-width: 72px;
  flex-shrink: 0;
}

.vc-slider-value {
  font-size: 11px;
  font-weight: 500;
  color: #a1a1aa;
  min-width: 56px;
  text-align: right;
  flex-shrink: 0;
}

.vc-slider-row input[type="range"] {
  flex: 1;
  height: 4px;
  -webkit-appearance: none;
  appearance: none;
  background: #3f3f46;
  border-radius: 2px;
  outline: none;
  cursor: pointer;
}

.vc-slider-row input[type="range"]::-webkit-slider-thumb {
  -webkit-appearance: none;
  appearance: none;
  width: 14px;
  height: 14px;
  border-radius: 50%;
  background: #8b5cf6;
  border: 2px solid #101014;
  box-shadow: 0 0 6px rgba(139, 92, 246, 0.4);
  cursor: pointer;
  transition: box-shadow 0.15s;
}

.vc-slider-row input[type="range"]::-webkit-slider-thumb:hover {
  box-shadow: 0 0 12px rgba(139, 92, 246, 0.6);
}

.vc-slider-row input[type="range"]::-moz-range-thumb {
  width: 14px;
  height: 14px;
  border-radius: 50%;
  background: #8b5cf6;
  border: 2px solid #101014;
  box-shadow: 0 0 6px rgba(139, 92, 246, 0.4);
  cursor: pointer;
}
"""

# The JS file includes the timing fix: API wrapping runs IMMEDIATELY at script
# load (polling every 10ms) instead of inside init() (which waits 500ms).
# This beats React's pywebview poller (50ms) and ensures all 9 speakers appear.
VOICE_CONTROLS_JS = """\
// Voice Creator Pro - Speed & Enthusiasm Controls + Expanded Speakers
(function () {
  'use strict';

  var OVERRIDE_PORT = 19876;
  var ENTHUSIASM_LABELS = ['Flat', 'Calm', 'Normal', 'Energetic', 'Intense'];
  var ENTHUSIASM_TEMPS = [0.3, 0.6, 0.9, 1.1, 1.3];
  var TTS_INSTRUCTIONS = [
    'Speak in a flat, monotone voice. ',
    'Speak in a calm, subdued tone. ',
    '',
    'Speak with energy and enthusiasm. ',
    'Speak with intense excitement and passion. '
  ];

  // --- localStorage persistence ---
  function loadVal(key, fallback) {
    try { var v = localStorage.getItem(key); return v !== null ? parseFloat(v) : fallback; }
    catch (e) { return fallback; }
  }
  function saveVal(key, v) {
    try { localStorage.setItem(key, String(v)); } catch (e) {}
  }

  var speedVal = loadVal('vc-speed', 1.0);
  var enthusiasmVal = loadVal('vc-enthusiasm', 3);

  // --- Send temperature override to backend ---
  function sendTemperatureOverride(sliderValue) {
    var temp = ENTHUSIASM_TEMPS[sliderValue - 1];
    fetch('http://127.0.0.1:' + OVERRIDE_PORT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ temperature: temp })
    }).catch(function () {});
  }

  // --- All Qwen3-TTS speakers (cross-lingual capable) ---
  var ALL_SPEAKERS = [
    { id: 'Vivian',   name: 'Vivian',   description: 'Bright, slightly edgy young female voice' },
    { id: 'Serena',   name: 'Serena',   description: 'Warm, gentle young female voice with a soft tone' },
    { id: 'Ryan',     name: 'Ryan',     description: 'Dynamic male voice with strong rhythmic drive' },
    { id: 'Aiden',    name: 'Aiden',    description: 'Sunny American male voice with a clear midrange' },
    { id: 'Ono_Anna', name: 'Ono Anna', description: 'Playful female voice with a light, nimble timbre' },
    { id: 'Sohee',    name: 'Sohee',    description: 'Warm female voice with rich emotion' },
    { id: 'Uncle_Fu', name: 'Uncle Fu', description: 'Seasoned male voice with a low, mellow timbre' },
    { id: 'Dylan',    name: 'Dylan',    description: 'Youthful male voice with a clear, natural timbre' },
    { id: 'Eric',     name: 'Eric',     description: 'Lively male voice with a slightly husky brightness' }
  ];

  // =======================================================================
  // TIMING FIX: Wrap API methods IMMEDIATELY at script load.
  // React's pywebview poller runs at 50ms intervals. Our script tag is a
  // regular (non-module) script that executes before the deferred module
  // bundle. By polling at 10ms we guarantee the wrapper is in place before
  // React ever calls get_tts_speakers().
  // =======================================================================
  (function earlyWrap() {
    var interval = setInterval(function () {
      if (!window.pywebview || !window.pywebview.api) return;
      var api = window.pywebview.api;
      if (api._vc_wrapped) { clearInterval(interval); return; }

      // Wrap get_tts_speakers -- inject all 9 speakers
      var origGetSpeakers = api.get_tts_speakers;
      if (typeof origGetSpeakers === 'function') {
        api.get_tts_speakers = function () {
          return origGetSpeakers.apply(api, arguments).then(function (result) {
            var speakers = (result && result.speakers) ? result.speakers : [];
            var existingIds = {};
            for (var i = 0; i < speakers.length; i++) existingIds[speakers[i].id] = true;
            for (var j = 0; j < ALL_SPEAKERS.length; j++) {
              if (!existingIds[ALL_SPEAKERS[j].id]) speakers.push(ALL_SPEAKERS[j]);
            }
            return { speakers: speakers };
          });
        };
      }

      // Wrap start_tts_generation -- prepend enthusiasm instruction
      var origTTS = api.start_tts_generation;
      if (typeof origTTS === 'function') {
        api.start_tts_generation = function () {
          var args = Array.prototype.slice.call(arguments);
          var idx = Math.round(enthusiasmVal) - 1;
          var prefix = TTS_INSTRUCTIONS[idx];
          if (prefix && args.length > 0 && typeof args[0] === 'string') {
            args[0] = prefix + args[0];
          }
          return origTTS.apply(api, args);
        };
      }

      api._vc_wrapped = true;
      clearInterval(interval);
    }, 10);
  })();

  // --- Detect current mode ---
  function getCurrentMode() {
    var tabs = document.querySelectorAll('[class*="bg-gradient"]');
    for (var i = 0; i < tabs.length; i++) {
      var text = (tabs[i].textContent || '').trim();
      if (text === 'Voice Clone') return 'clone';
      if (text === 'Text to Speech') return 'tts';
      if (text === 'Voice Design') return 'design';
    }
    var allBtns = document.querySelectorAll('button');
    for (var j = 0; j < allBtns.length; j++) {
      var btn = allBtns[j];
      var t = (btn.textContent || '').trim();
      var hasGradient = btn.className && btn.className.indexOf('gradient') !== -1;
      if (hasGradient) {
        if (t === 'Voice Clone') return 'clone';
        if (t === 'Text to Speech') return 'tts';
        if (t === 'Voice Design') return 'design';
      }
    }
    return 'clone';
  }

  // --- Build slider UI ---
  function buildControls() {
    var container = document.createElement('div');
    container.id = 'vc-voice-controls';

    // Speed slider
    var speedRow = document.createElement('div');
    speedRow.className = 'vc-slider-row';
    var speedLabel = document.createElement('span');
    speedLabel.className = 'vc-slider-label';
    speedLabel.textContent = 'Speed';
    var speedInput = document.createElement('input');
    speedInput.type = 'range';
    speedInput.min = '0.5';
    speedInput.max = '2.0';
    speedInput.step = '0.05';
    speedInput.value = String(speedVal);
    var speedValue = document.createElement('span');
    speedValue.className = 'vc-slider-value';
    speedValue.textContent = speedVal.toFixed(2) + 'x';
    speedInput.addEventListener('input', function () {
      speedVal = parseFloat(speedInput.value);
      speedValue.textContent = speedVal.toFixed(2) + 'x';
      saveVal('vc-speed', speedVal);
      applySpeedToAudio();
    });
    speedRow.appendChild(speedLabel);
    speedRow.appendChild(speedInput);
    speedRow.appendChild(speedValue);

    // Enthusiasm slider
    var enthRow = document.createElement('div');
    enthRow.className = 'vc-slider-row';
    var enthLabel = document.createElement('span');
    enthLabel.className = 'vc-slider-label';
    enthLabel.textContent = 'Enthusiasm';
    var enthInput = document.createElement('input');
    enthInput.type = 'range';
    enthInput.min = '1';
    enthInput.max = '5';
    enthInput.step = '1';
    enthInput.value = String(Math.round(enthusiasmVal));
    var enthValue = document.createElement('span');
    enthValue.className = 'vc-slider-value';
    enthValue.textContent = ENTHUSIASM_LABELS[Math.round(enthusiasmVal) - 1];
    enthInput.addEventListener('input', function () {
      enthusiasmVal = parseInt(enthInput.value, 10);
      enthValue.textContent = ENTHUSIASM_LABELS[enthusiasmVal - 1];
      saveVal('vc-enthusiasm', enthusiasmVal);
      var mode = getCurrentMode();
      if (mode === 'clone' || mode === 'design') {
        sendTemperatureOverride(enthusiasmVal);
      }
    });
    enthRow.appendChild(enthLabel);
    enthRow.appendChild(enthInput);
    enthRow.appendChild(enthValue);

    container.appendChild(speedRow);
    container.appendChild(enthRow);
    return container;
  }

  // --- Apply playback speed to all audio elements ---
  function applySpeedToAudio() {
    var audios = document.querySelectorAll('audio');
    for (var i = 0; i < audios.length; i++) {
      audios[i].preservesPitch = true;
      audios[i].playbackRate = speedVal;
    }
  }

  // --- Observe for new audio elements ---
  function watchForAudio() {
    var obs = new MutationObserver(function () {
      applySpeedToAudio();
    });
    obs.observe(document.body, { childList: true, subtree: true });
  }

  // --- Inject controls into DOM ---
  function inject() {
    if (document.getElementById('vc-voice-controls')) return;

    var candidates = document.querySelectorAll('span, div, label');
    var qualityParent = null;
    for (var i = 0; i < candidates.length; i++) {
      var text = (candidates[i].textContent || '').trim();
      if (text === 'Quality' || text === 'Quality:') {
        qualityParent = candidates[i].closest('div[class]');
        if (qualityParent) {
          var section = qualityParent.parentElement;
          if (section) {
            qualityParent = section;
          }
        }
        break;
      }
    }

    var controls = buildControls();

    if (qualityParent && qualityParent.parentElement) {
      qualityParent.parentElement.insertBefore(controls, qualityParent.nextSibling);
    } else {
      var sidebar = document.querySelector('[class*="space-y-3"], [class*="space-y-4"]');
      if (sidebar) {
        sidebar.appendChild(controls);
      }
    }
  }

  // --- Wait for app to render, then inject DOM elements ---
  // NOTE: API wrapping already happened above (earlyWrap). init() only
  // handles DOM injection, audio speed, and the initial temperature push.
  function init() {
    inject();

    var obs = new MutationObserver(function () {
      if (!document.getElementById('vc-voice-controls')) {
        inject();
      }
    });
    obs.observe(document.body, { childList: true, subtree: true });

    watchForAudio();
    sendTemperatureOverride(Math.round(enthusiasmVal));
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    setTimeout(init, 500);
  }
})();
"""

# Content to inject into utils.py (HTTP override server + temperature hook)
UTILS_OVERRIDE_SERVER = '''\

# --- Voice Creator Pro: runtime generation overrides ---
_generation_overrides = {}

class _OverrideHandler(__import__('http.server', fromlist=['BaseHTTPRequestHandler']).BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        data = _json.loads(self.rfile.read(length)) if length else {}
        _generation_overrides.update(data)
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(_json.dumps(_generation_overrides).encode())
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    def log_message(self, *args): pass

try:
    import http.server as _http_server
    _override_server = _http_server.HTTPServer(('127.0.0.1', 19876), _OverrideHandler)
    threading.Thread(target=_override_server.serve_forever, daemon=True).start()
except OSError:
    pass
# --- End Voice Creator Pro override server ---
'''

UTILS_TEMP_INJECTION = """\
        # Voice Creator Pro: apply runtime temperature overrides
        if _generation_overrides.get("temperature") is not None:
            generation_config.temperature = float(_generation_overrides["temperature"])
"""


def fix_ui_sliders():
    """Patch utils.py with override server, write CSS/JS, patch index.html."""
    print("\n[Fix 3+4] Speed/Enthusiasm Sliders + Expanded Speaker List")
    all_ok = True

    # --- 3a: Patch utils.py ---
    if not os.path.isfile(UTILS_PY):
        status("utils.py", False, "file not found")
        all_ok = False
    else:
        with open(UTILS_PY, "r", encoding="utf-8") as f:
            content = f.read()

        server_marker = "_generation_overrides = {}"
        temp_marker = "Voice Creator Pro: apply runtime temperature overrides"

        if server_marker in content and temp_marker in content:
            status("utils.py", True, "already patched")
        else:
            backup(UTILS_PY)
            modified = False

            # Insert override server after the logger line
            if server_marker not in content:
                anchor = 'logger = logging.get_logger(__name__)\n'
                if anchor in content:
                    content = content.replace(anchor, anchor + UTILS_OVERRIDE_SERVER)
                    modified = True
                else:
                    status("utils.py override server", False, "anchor not found")
                    all_ok = False

            # Insert temperature injection in generate()
            if temp_marker not in content:
                gen_anchor = (
                    "        generation_config, model_kwargs = self._prepare_generation_config(\n"
                    "            generation_config, use_model_defaults, **kwargs\n"
                    "        )\n"
                )
                if gen_anchor in content:
                    content = content.replace(gen_anchor, gen_anchor + UTILS_TEMP_INJECTION)
                    modified = True
                else:
                    status("utils.py temperature injection", False, "anchor not found")
                    all_ok = False

            if modified:
                with open(UTILS_PY, "w", encoding="utf-8") as f:
                    f.write(content)
                status("utils.py", True, "patched")

    # --- 3b: Write CSS ---
    os.makedirs(ASSETS_DIR, exist_ok=True)
    try:
        with open(CSS_FILE, "w", encoding="utf-8") as f:
            f.write(VOICE_CONTROLS_CSS)
        status("voice-controls.css", True, "written")
    except Exception as exc:
        status("voice-controls.css", False, str(exc))
        all_ok = False

    # --- 3c: Write JS (includes Fix 4 timing fix) ---
    try:
        with open(JS_FILE, "w", encoding="utf-8") as f:
            f.write(VOICE_CONTROLS_JS)
        status("voice-controls.js", True, "written (with timing fix)")
    except Exception as exc:
        status("voice-controls.js", False, str(exc))
        all_ok = False

    # --- 3d: Patch index.html ---
    if not os.path.isfile(INDEX_HTML):
        status("index.html", False, "file not found")
        all_ok = False
    else:
        with open(INDEX_HTML, "r", encoding="utf-8") as f:
            html = f.read()

        css_tag = '<link rel="stylesheet" href="./assets/voice-controls.css">'
        js_tag = '<script src="./assets/voice-controls.js"></script>'

        if css_tag in html and js_tag in html:
            status("index.html", True, "already patched")
        else:
            backup(INDEX_HTML)
            # Insert before </body>
            inject = ""
            if css_tag not in html:
                inject += "    " + css_tag + "\n"
            if js_tag not in html:
                inject += "    " + js_tag + "\n"

            if inject and "</body>" in html:
                html = html.replace("</body>", inject + "  </body>")
                with open(INDEX_HTML, "w", encoding="utf-8") as f:
                    f.write(html)
                status("index.html", True, "patched")
            elif not inject:
                status("index.html", True, "already has tags")
            else:
                status("index.html", False, "</body> tag not found")
                all_ok = False

    return all_ok


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("  Voice Creator Pro - Comprehensive Patcher")
    print("=" * 60)

    if not check_admin():
        print("\n  WARNING: Not running as Administrator.")
        print("  Fix 1 (Long Paths) requires elevation.")
        print("  Right-click Command Prompt -> Run as Administrator")
        print()

    results = {}
    results["Long Paths"] = fix_long_paths()
    results["GPU/Triton"] = fix_gpu_triton()
    results["Sliders+Speakers"] = fix_ui_sliders()

    print("\n" + "=" * 60)
    print("  Summary")
    print("=" * 60)
    all_ok = True
    for name, ok in results.items():
        tag = "OK" if ok else "FAIL"
        print(f"  [{tag}] {name}")
        if not ok:
            all_ok = False

    if all_ok:
        print("\n  All fixes applied successfully.")
        print("  Reboot recommended (required for Long Paths).")
        print("  Then launch Voice Creator Pro and verify:")
        print("    - TTS mode shows all 9 speakers in dropdown")
        print("    - Speed and Enthusiasm sliders appear below Quality")
        print("    - Speech generation uses CUDA without Triton errors")
    else:
        print("\n  Some fixes had issues. Review the output above.")

    print()


if __name__ == "__main__":
    main()

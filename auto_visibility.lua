local obs = obslua
local ffi = require("ffi")

-- Windows API関数を定義
ffi.cdef [[
  typedef void* HWND;
  typedef unsigned long DWORD;
  typedef int BOOL;

  HWND GetForegroundWindow();
  int GetWindowTextW(HWND hWnd, wchar_t* lpString, int nMaxCount);
  DWORD GetWindowThreadProcessId(HWND hWnd, DWORD* lpdwProcessId);

  typedef void* HANDLE;
  HANDLE OpenProcess(DWORD dwDesiredAccess, int bInheritHandle, DWORD dwProcessId);
  int CloseHandle(HANDLE hObject);
  DWORD GetModuleBaseNameW(HANDLE hProcess, void* hModule, wchar_t* lpBaseName, DWORD nSize);

  int WideCharToMultiByte(
    unsigned int CodePage,
    unsigned long dwFlags,
    const wchar_t* lpWideCharStr,
    int cchWideChar,
    char* lpMultiByteStr,
    int cbMultiByte,
    const char* lpDefaultChar,
    int* lpUsedDefaultChar
  );
]]

local user32 = ffi.load("user32")
local kernel32 = ffi.load("kernel32")
local psapi = ffi.load("psapi")

-- 定数定義
local CP_UTF8 = 65001
local PROCESS_QUERY_INFORMATION = 0x0400
local PROCESS_VM_READ = 0x0010

-- グローバル変数
local monitored_sources = {}
local interval_msec = 1000
local debug_mode = false
local source_1 = ""
local source_2 = ""
local source_3 = ""
local source_4 = ""
local source_5 = ""

-------------------------------------------------------------------
-- アクティブウィンドウのプロセス名を取得する
-------------------------------------------------------------------
local function get_active_process_name()
  local hwnd = user32.GetForegroundWindow()
  if hwnd == nil then
    return ""
  end

  -- プロセスIDを取得
  local process_id = ffi.new("DWORD[1]")
  user32.GetWindowThreadProcessId(hwnd, process_id)

  if process_id[0] == 0 then
    return ""
  end

  -- プロセスハンドルを開く
  local process_handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION + PROCESS_VM_READ, 0, process_id[0])
  if process_handle == nil then
    return ""
  end

  local buffer_size = 256
  local buffer = ffi.new("wchar_t[?]", buffer_size)
  local length = psapi.GetModuleBaseNameW(process_handle, nil, buffer, buffer_size)

  local result = ""
  if length > 0 then
    -- まず必要なUTF-8バイト数を取得
    local utf8_size = kernel32.WideCharToMultiByte(CP_UTF8, 0, buffer, length, nil, 0, nil, nil)
    if utf8_size > 0 then
      -- UTF-8バッファを確保
      local utf8_buffer = ffi.new("char[?]", utf8_size + 1)
      -- 実際のUTF-16からUTF-8への変換を実行
      local result_size = kernel32.WideCharToMultiByte(CP_UTF8, 0, buffer, length, utf8_buffer, utf8_size, nil, nil)
      if result_size > 0 then
        result = ffi.string(utf8_buffer, result_size)
      end
    end
  end

  -- プロセスハンドルを閉じる
  kernel32.CloseHandle(process_handle)

  return result
end

-------------------------------------------------------------------
-- ウィンドウ識別子からプロセス名部分を抽出する
-- 識別子は "タイトル:クラス:実行ファイル名" という形式
-------------------------------------------------------------------
local function get_process_name_from_identifier(identifier)
  if not identifier or identifier == "" then
    return nil
  end
  -- 最後のコロン ":" の後の部分がプロセス名（実行ファイル名）
  local last_colon_pos = identifier:match(".*:()")
  if last_colon_pos then
    return string.sub(identifier, last_colon_pos)
  end
  return identifier
end

-------------------------------------------------------------------
-- UIで設定された内容をグローバル変数に反映する
-------------------------------------------------------------------
function script_update(settings)
  local new_monitored = {}

  -- 各ソース設定を読み込む
  source_1 = obs.obs_data_get_string(settings, "source_1")
  source_2 = obs.obs_data_get_string(settings, "source_2")
  source_3 = obs.obs_data_get_string(settings, "source_3")
  source_4 = obs.obs_data_get_string(settings, "source_4")
  source_5 = obs.obs_data_get_string(settings, "source_5")

  local source_names = { source_1, source_2, source_3, source_4, source_5 }

  for _, source_name in ipairs(source_names) do
    if source_name and source_name ~= "" then
      local source = obs.obs_get_source_by_name(source_name)
      if source then
        local source_settings = obs.obs_source_get_settings(source)
        -- "window" または "capture_window" キーからウィンドウ識別子を取得
        local window_identifier = obs.obs_data_get_string(source_settings, "window")
        if window_identifier == "" then
          window_identifier = obs.obs_data_get_string(source_settings, "capture_window") -- ゲームキャプチャ用のキー
        end

        local process_name_to_match = get_process_name_from_identifier(window_identifier)

        if process_name_to_match then
          table.insert(new_monitored, {
            source_name = source_name,
            target_process_name = process_name_to_match
          })
        end

        obs.obs_data_release(source_settings)
        obs.obs_source_release(source)
      end
    end
  end

  monitored_sources = new_monitored
  interval_msec = obs.obs_data_get_int(settings, "interval")
  debug_mode = obs.obs_data_get_bool(settings, "debug_mode")
end

-------------------------------------------------------------------
-- 定期的にアクティブウィンドウをチェックして表示を切り替える
-------------------------------------------------------------------
local function check_active_window()
  local active_process_name = get_active_process_name()

  for _, item in ipairs(monitored_sources) do
    local source = obs.obs_get_source_by_name(item.source_name)
    if source then
      local scenes = obs.obs_frontend_get_scenes()
      if scenes ~= nil then
        for _, scene_source in ipairs(scenes) do
          local scene = obs.obs_scene_from_source(scene_source)
          local scene_item = obs.obs_scene_find_source(scene, item.source_name)
          if scene_item then
            -- デバッグログ: プロセス名比較の情報を出力（デバッグモード時のみ）
            local match_found = active_process_name ~= "" and active_process_name == item.target_process_name
            if debug_mode then
              obs.script_log(obs.LOG_INFO, string.format("デバッグ: アクティブプロセス='%s', 監視対象プロセス='%s', マッチ=%s",
                active_process_name or "", item.target_process_name or "", tostring(match_found)))
            end

            -- アクティブウィンドウのプロセス名が、キャプチャ対象のプロセス名と一致するか
            if match_found then
              obs.obs_sceneitem_set_visible(scene_item, true)
            else
              obs.obs_sceneitem_set_visible(scene_item, false)
            end
          end
        end
      end
      obs.source_list_release(scenes)
      obs.obs_source_release(source)
    end
  end
end

-------------------------------------------------------------------
-- ソースリストを作成するヘルパー関数
-------------------------------------------------------------------
local function add_sources_to_list(property, sources)
  if sources ~= nil then
    for _, source in ipairs(sources) do
      local source_id = obs.obs_source_get_id(source)
      if source_id == "window_capture" or source_id == "game_capture" then
        local name = obs.obs_source_get_name(source)
        obs.obs_property_list_add_string(property, name, name)
      end
    end
  end
end

-------------------------------------------------------------------
-- スクリプトのUIを定義する
-------------------------------------------------------------------
function script_properties()
  local props = obs.obs_properties_create()

  obs.obs_properties_add_text(props, "info", "ウィンドウキャプチャ/ゲームキャプチャソースを5つまで登録できます。", obs.OBS_TEXT_INFO)

  -- ソースのリストを取得
  local sources = obs.obs_enum_sources()

  -- 5つのソース選択ドロップダウンを作成
  local p1 = obs.obs_properties_add_list(props, "source_1", "ソース 1", obs.OBS_COMBO_TYPE_LIST, obs
    .OBS_COMBO_FORMAT_STRING)
  obs.obs_property_list_add_string(p1, "なし", "")
  add_sources_to_list(p1, sources)

  local p2 = obs.obs_properties_add_list(props, "source_2", "ソース 2", obs.OBS_COMBO_TYPE_LIST, obs
    .OBS_COMBO_FORMAT_STRING)
  obs.obs_property_list_add_string(p2, "なし", "")
  add_sources_to_list(p2, sources)

  local p3 = obs.obs_properties_add_list(props, "source_3", "ソース 3", obs.OBS_COMBO_TYPE_LIST, obs
    .OBS_COMBO_FORMAT_STRING)
  obs.obs_property_list_add_string(p3, "なし", "")
  add_sources_to_list(p3, sources)

  local p4 = obs.obs_properties_add_list(props, "source_4", "ソース 4", obs.OBS_COMBO_TYPE_LIST, obs
    .OBS_COMBO_FORMAT_STRING)
  obs.obs_property_list_add_string(p4, "なし", "")
  add_sources_to_list(p4, sources)

  local p5 = obs.obs_properties_add_list(props, "source_5", "ソース 5", obs.OBS_COMBO_TYPE_LIST, obs
    .OBS_COMBO_FORMAT_STRING)
  obs.obs_property_list_add_string(p5, "なし", "")
  add_sources_to_list(p5, sources)

  -- ソースリストの解放
  obs.source_list_release(sources)

  obs.obs_properties_add_int(props, "interval", "監視間隔(ms)", 500, 10000, 100)

  -- デバッグモードのチェックボックス
  obs.obs_properties_add_bool(props, "debug_mode", "debug")

  return props
end

-------------------------------------------------------------------
-- スクリプトの初期化処理
-------------------------------------------------------------------
function script_load(settings)
  script_update(settings)
  obs.timer_add(check_active_window, interval_msec)
end

function script_defaults(settings)
  obs.obs_data_set_default_int(settings, "interval", 1000)
  obs.obs_data_set_default_bool(settings, "debug_mode", false)
end

function script_description()
  return "ウィンドウキャプチャ/ゲームキャプチャソースを選択し、それぞれのプロセスがアクティブな時に自動で表示します。\n最大5つのソースを監視できます。プロセス名で判定するため、ウィンドウタイトルが変わっても動作します。"
end

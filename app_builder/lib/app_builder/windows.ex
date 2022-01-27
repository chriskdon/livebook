defmodule AppBuilder.Windows do
  @moduledoc false

  import AppBuilder.Utils
  require EEx

  @doc """
  Creates a Windows installer.
  """
  def build_windows_installer(release, options) do
    options =
      Keyword.validate!(options, [
        :name,
        :url_schemes,
        :logo_path
      ])

    app_name = Keyword.fetch!(options, :name)

    tmp_dir = "tmp"
    File.mkdir_p!(tmp_dir)
    File.rm_rf!("tmp/rel")

    logo_path = options[:logo_path] || Application.app_dir(:wx, "examples/demo/erlang.png")
    app_icon_path = Path.join(tmp_dir, "app_icon.ico")
    copy_image(logo_path, app_icon_path)

    File.rename!(release.path, Path.join(tmp_dir, "rel"))
    erts_dir = Path.join([tmp_dir, "rel", "erts-#{:erlang.system_info(:version)}"])

    rcedit_path = ensure_rcedit(tmp_dir)

    cmd!(rcedit_path, [
      "--set-icon",
      app_icon_path,
      Path.join([erts_dir, "bin", "erl.exe"])
    ])

    File.write!(Path.join(tmp_dir, "#{app_name}.vbs"), launcher(release))
    nsi_path = Path.join(tmp_dir, "#{app_name}.nsi")
    File.write!(nsi_path, nsi(options))
    cmd!("makensis", [nsi_path])

    File.rename!(
      Path.join(tmp_dir, "#{app_name}Install.exe"),
      Path.join([Mix.Project.build_path(), "rel", "#{app_name}Install.exe"])
    )

    release
  end

  code = """
  <%
  app_name = Keyword.fetch!(options, :name)
  url_schemes = Keyword.get(options, :url_schemes, [])
  %>
  !include "MUI2.nsh"

  ;--------------------------------
  ;General

  Name "<%= app_name %>"
  OutFile "<%= app_name %>Install.exe"
  Unicode True
  InstallDir "$LOCALAPPDATA\\<%= app_name %>"
  ; need admin for registering URL scheme, otherwise user would suffice.
  RequestExecutionLevel admin

  ;--------------------------------
  ;Interface Settings

  !define MUI_ABORTWARNING

  ;--------------------------------
  ;Pages

  ;!insertmacro MUI_PAGE_COMPONENTS
  !define MUI_ICON "app_icon.ico"
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES

  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES

  ;--------------------------------
  ;Languages

  !insertmacro MUI_LANGUAGE "English"

  ;--------------------------------
  ;Installer Sections

  Section "Dummy Section" SecDummy
    SetOutPath "$INSTDIR"
    ;CopyFiles "${NSISDIR}\\<%= app_name %>Install.exe" "$INSTDIR\\<%= app_name %>.exe"
    File /r rel rel
    File "<%= app_name %>.vbs"
    WriteUninstaller "$INSTDIR\\<%= app_name %>Uninstall.exe"

  <%= for url_scheme <- url_schemes do %>
    DetailPrint "Register <%= url_scheme %> URL Handler"
    DeleteRegKey HKCR "<%= url_scheme %>"
    WriteRegStr  HKCR "<%= url_scheme %>" "" "URL:<%= url_scheme %> Protocol"
    WriteRegStr  HKCR "<%= url_scheme %>" "URL Protocol" ""
    WriteRegStr  HKCR "<%= url_scheme %>\\shell" "" ""
    WriteRegStr  HKCR "<%= url_scheme %>\\shell\\open" "" ""
    WriteRegStr  HKCR "<%= url_scheme %>\\shell\\open\\command" "" '$WINDIR\\system32\\wscript.exe "$INSTDIR\\<%= app_name %>.vbs" "%1"'
  <% end %>
  SectionEnd

  Section "Desktop Shortcut" SectionX
    CreateShortCut "$DESKTOP\\<%= app_name %>.lnk" "$INSTDIR\\<%= app_name %>.vbs"
  SectionEnd

  Section "Uninstall"
    Delete "$INSTDIR\\<%= app_name %>Install.exe"
    Delete "$INSTDIR\\<%= app_name %>Uninstall.exe"
    Delete "$DESKTOP\\<%= app_name %>.lnk"
    RMDir "$INSTDIR"
  SectionEnd
  """

  EEx.function_from_string(:defp, :nsi, code, [:options], trim: true)

  code = """
  ' This avoids a flashing cmd window when launching the bat file
  strPath = Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullName) - Len(Wscript.ScriptName)) & "rel\\bin\\<%= release.name %>.bat"
  ' Debug: MsgBox(strPath)

  Dim Args()
  ReDim Args(WScript.Arguments.Count - 1)

  For i = 0 To WScript.Arguments.Count - 1
     Args(i) = \"""" & WScript.Arguments(i) & \""""
  Next

  Set WshShell = CreateObject("WScript.Shell" )
  WshShell.Run \"""" & strPath & \""" start -- " & Join(Args), 0
  Set WshShell = Nothing
  """

  EEx.function_from_string(:defp, :launcher, code, [:release], trim: true)

  defp ensure_rcedit(tmp_dir) do
    rcedit_path = Path.join([File.cwd!(), tmp_dir, "rcedit"])

    unless File.exists?(rcedit_path) do
      cmd!("curl", [
        "-L",
        "https://github.com/electron/rcedit/releases/download/v1.1.1/rcedit-x64.exe",
        "-o",
        rcedit_path
      ])
    end

    rcedit_path
  end

  defp copy_image(src_path, dest_path) do
    if Path.extname(src_path) == ".ico" do
      File.cp!(src_path, dest_path)
    else
      cmd!("magick", [
        src_path,
        "-background",
        "none",
        "-resize",
        "128x128",
        dest_path
      ])
    end
  end
end

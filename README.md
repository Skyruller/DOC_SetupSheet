# DOC SetupSheet

### PowerMILL Plugin for Automatic Tech Card Generation

[![PowerMILL](https://img.shields.io/badge/PowerMILL-2018%2B-blue)](https://www.autodesk.com/products/powermill)\
[![.NET](https://img.shields.io/badge/.NET-4.0-green)](https://dotnet.microsoft.com/)\
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

------------------------------------------------------------------------

## 📌 About

**DOC SetupSheet** --- плагин для Autodesk PowerMILL, который
автоматически генерирует технологические карты (DOCX) на основе данных
текущего проекта.

Плагин извлекает информацию из NC-программ, инструментов и проекта,
формирует документ Word по шаблону и сохраняет его в папке проекта.

------------------------------------------------------------------------

## 🎯 Features

-   ✅ Генерация Word-документа из шаблона (.docx)
-   ✅ Автоматический сбор NC-программ
-   ✅ Список инструментов из выбранных NC
-   ✅ Подсчёт общего времени обработки
-   ✅ Вставка скриншота из PowerMILL
-   ✅ Настраиваемый шаблон документа
-   ✅ Поддержка сетевых путей
-   ✅ Совместимость PowerMILL 2018--2024

------------------------------------------------------------------------

## 📦 Requirements

  Component            Version
  -------------------- -------------------
  Autodesk PowerMILL   2018+
  .NET Framework       4.0+
  Microsoft Word       2010+
  Visual Studio        2015+ (for build)

------------------------------------------------------------------------

## 🛠 Installation

### 1️⃣ Build

Open solution in Visual Studio and run:

Build → Rebuild Solution

After build you will get:

SetupSheet.dll

------------------------------------------------------------------------

### 2️⃣ Copy Files

Create plugin folder, for example:

C:`\Program `{=tex}Files`\Autodesk`{=tex}`\DOC`{=tex}\_SetupSheet\

Copy:

-   SetupSheet.dll\
-   TechCard_Template.docx\
-   Icons folder

------------------------------------------------------------------------

### 3️⃣ Register COM

Run command prompt as Administrator:

C:`\WINDOWS`{=tex}`\Microsoft`{=tex}.NET`\Framework64`{=tex}`\v4`{=tex}.0.30319`\regasm`{=tex}.exe
"C:`\Program `{=tex}Files`\Autodesk`{=tex}`\DOC`{=tex}\_SetupSheet`\SetupSheet`{=tex}.dll"
/register /codebase

reg.exe ADD
"HKCR`\CLSID`{=tex}{8C96851C-7A01-4389-8FBF-22C3DC7B09FD}`\Implemented `{=tex}Categories{311b0135-1826-4a8c-98de-f313289f815e}"
/reg:64 /f

------------------------------------------------------------------------

### 4️⃣ Enable in PowerMILL

1.  File → Options → Manage Installed Plugins\
2.  Find **DOC SetupSheet**\
3.  Click **Enable**\
4.  Restart PowerMILL

------------------------------------------------------------------------

## 📄 How It Works

1.  Select NC programs\
2.  Plugin collects:
    -   tools
    -   machining time
    -   program order
    -   screenshot
3.  Word document is generated from template\
4.  File is saved to:

\[ProjectPath\]`\Excel`{=tex}\_Setupsheet`\Техкарта`{=tex}.docx

------------------------------------------------------------------------

## 🧩 Template Placeholders

  Placeholder     Description
  --------------- ------------------
  {article}       Article
  {material}      Material
  {machines}      Machine
  {stockSize}     Stock size
  {projectPath}   Project path
  {ncList}        NC programs
  {toolsList}     Tools list
  {time}          Machining time
  {comments}      Comments
  {screenshot}    Screenshot image

⚠ Placeholders must match exactly (case-sensitive).

------------------------------------------------------------------------

## 📁 Project Structure

SetupSheet/ ├── ToolSheet.cs\
├── ToolSheetPaneWPF.xaml\
├── ToolSheetPaneWPF.xaml.cs\
├── DocumentGenerator.cs\
├── Options.cs\
├── TechCard_Template.docx\
├── Icons/\
└── README.md

------------------------------------------------------------------------

## 🛠 Troubleshooting

### Time or Tools not inserted

-   Check that `{time}` and `{toolsList}` exist in template\
-   Ensure no extra spaces inside braces\
-   Verify data is passed from ToolSheetPaneWPF

### Word shows "Save As" dialog

-   File is opened in Word\
-   Path is not accessible\
-   Network drive permissions issue

------------------------------------------------------------------------

## 📜 License

MIT License.

------------------------------------------------------------------------

## 👤 Author

Developed by **Skyruller**\
Based on PowerMILL API examples.

------------------------------------------------------------------------

Version: 1.0.0\
PowerMILL: 2018--2024\
GUID: {8C96851C-7A01-4389-8FBF-22C3DC7B09FD}

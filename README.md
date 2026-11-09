# 📚 Документація VBA-системи обліку МТЗ військової частини

> **Версія коду:** ~3.9.26 (за мітками у коментарях)  
> **Платформа:** Microsoft Excel (VBA)  
> **Призначення:** Автоматизований облік матеріально-технічного забезпечення (МТЗ) підрозділів ЗСУ

---

## 🏛️ Загальний опис системи

Система є **Excel-застосунком** для автоматизованого обліку матеріально-технічного забезпечення (МТЗ) військової частини. Автоматизує формування первинних облікових документів відповідно до наказів Міністерства оборони України.

### Що автоматизує система:
- 📥 **Імпорт** заповнених облікових документів у центральний журнал (`AllData`)
- 📄 **Формування** первинних документів (накладних, актів, відомостей)
- 📊 **Звітність** — залишки ТМЦ за підрозділами, особами, номенклатурою
- 🔍 **Перевірку** якості даних у журналі

### Нормативна база:
| Наказ МОУ | Форми |
|-----------|-------|
| Наказ № 81 | Акт списання ф.3, Акт зміни стану ф.14, Акт якісного стану ф.1 |
| Наказ № 440 | Накладна ф.25, Відомість ф.57, Акти ОЗ ф.23/24, Акт розміщення ф.45, Звітна ф.18 |
| Наказ № 431 | Акт списання, Накладна |
| Наказ № 818 | Акт прийому/передачі ОЗ |

---

## 📁 Структура документації

| Файл | Зміст |
|------|-------|
| **README.md** | Цей файл — загальний огляд системи |
| **[01_Architecture.md](01_Architecture.md)** | Архітектура: схема аркушів, взаємодія модулів, точки входу |
| **[02_Tables_Structure.md](02_Tables_Structure.md)** | Структура таблиць: AllData, залишки, Named Ranges |
| **[03_Core_Modules.md](03_Core_Modules.md)** | Базові модулі: functions, checkErrors, ManualAdd, SumInWords, MoveForUnit |
| **[04_UniversalImport.md](04_UniversalImport.md)** | Модуль UniversalImport — масовий імпорт документів |
| **[05_Residues_Reports.md](05_Residues_Reports.md)** | Модулі залишків: Residuse, ResiduseOneTroop/Good/People |
| **[06_Documents_WOff_Invoice.md](06_Documents_WOff_Invoice.md)** | Акти списання та накладні: WOff431/81/19, invoice431, Invoice440_25, payroll440_57 |
| **[07_Documents_Acts.md](07_Documents_Acts.md)** | Акти ОЗ та заявки: ActRTFA440_23/24, ActRTFA818, ActChangeState/QState, ActPlacement, RequestM |
| **[08_Reports_AddOns.md](08_Reports_AddOns.md)** | Звіти-додатки: Додаток1, Відомість14, Додаток47, Звітна440_18 |
| **[09_Reports_Inventory.md](09_Reports_Inventory.md)** | Інвентаризаційні описи оборотних та необоротних активів |
| **[10_Classes_Forms.md](10_Classes_Forms.md)** | Класи аркушів (wsAllData, wsResiduse2, Аркуш*) та форми (UserForm1, ProgressBar) |

---

## 🗂️ Структура файлів VBA проекту

### Класи аркушів (.cls)
```
wsAllData.cls          — Клас аркуша «Дані» (головний журнал AllData)
wsResiduse2.cls        — Клас аркуша «Залишки 2» (пул залишків)
Аркуш5.cls             — Клас документного аркуша (накладна/відомість)
Аркуш11.cls            — Клас документного аркуша
Аркуш18.cls            — Клас аркуша «Заявки»
Аркуш20.cls            — Клас документного аркуша
clsSearchableDropdown.cls — Компонент пошукового ListBox (Paul Kelly v2.2)
```

### Форми (.frm)
```
UserForm1.frm          — Форма пошукового вибору товару/підрозділу
ProgressBar.frm        — Форма прогресбару для тривалих операцій
```

### Службові модулі (.bas)
```
Module1.bas            — Утиліта: експорт VBA компонентів у файли
RibbonCallbacks.bas    — Обробники кнопок кастомної стрічки Ribbon
```

### Базові/допоміжні модулі (.bas)
```
functions.bas          — Загальні утиліти (7 функцій/процедур)
checkErrors.bas        — Перевірка якості даних в AllData
ManualAdd.bas          — Ручне введення запису через InputBox
SumInWords.bas         — Конвертація числа у текст прописом (укр.)
MoveForUnit.bas        — Звіт руху ТМЦ по підрозділу
```

### Модуль імпорту (.bas)
```
UniversalImport.bas    — Масовий імпорт документів (11 типів), 104 КБ
```

### Модулі залишків і звітів (.bas)
```
Residuse.bas           — Зведені залишки по підрозділах (PivotTable+SUMIFS)
Reiduse2.bas           — Деталізовані залишки для всіх підрозділів (Dictionary)
ResiduseOneTroop.bas   — Залишки одного підрозділу v1
ResiduseOneTroop2.bas  — Залишки одного підрозділу v2 (з датами)
ResiduseOneGood.bas    — Залишки одного товару по підрозділах
ResiduseOnePeople.bas  — Залишки по МВО (матеріально-відповідальній особі)
ReportingMove.bas      — Звітна відомість руху 440_18
add_1_one_Troop.bas    — Додаток 1 v1
add_1_one_Troop2.bas   — Додаток 1 v2 (по МВО, масовий), 69 КБ
Add14.bas              — Відомість №14 (всі підрозділи)
Add14OneTroop.bas      — Відомість №14 (один підрозділ)
Add47.bas              — Додаток 47
add47PNP.bas           — Додаток 47 ПН-підрозділ
Inventory.bas          — Інвентаризаційний опис (оборотні активи)
InvertoryNonCurrentAssets.bas — Інвентаризаційний опис (ОЗ, НМА)
```

### Модулі первинних документів (.bas)
```
WOff431.bas            — Акт списання 431
WOff81.bas             — Акт списання 81
WOffLaunch81_19.bas    — Акт введення/виведення з експлуатації №19 (81)
invoice431.bas         — Накладна 431
Invoice440_25.bas      — Накладна 440 ф.25
payroll440_57.bas      — Відомість(а)ність №57
ActRTFA440_23.bas      — Акт прийому/передачі ОЗ 440 ф.23
ActRTFA440_24.bas      — Акт передачі ОЗ 440 ф.24
ActRTFA818.bas         — Акт ОЗ наказ 818
ActChangeState81_14.bas — Акт зміни стану майна №14 (81)
ActQState81_1.bas      — Акт якісного стану майна №1 (81)
ActPlacement440_45.bas — Акт прийому (розміщення) №45 (440)
RequestM.bas           — Заявка на отримання майна
```

---

## 🎯 Ключові точки входу

| Точка входу | Модуль | Як викликається |
|-------------|--------|-----------------|
| `UnivImport()` | UniversalImport.bas | Кнопка «Імпорт» на стрічці Ribbon |
| `ManualAdding()` | ManualAdd.bas | Кнопка «Ручне введення» на стрічці Ribbon |
| `Worksheet_BeforeRightClick` | wsAllData.cls, wsResiduse2.cls, Аркуш*.cls | Правий клік на аркуші |
| `addPos*()` | Всі документні модулі | Кнопки на аркушах документів |
| `Save*()` | Всі документні модулі | Кнопки «Зберегти» на аркушах |
| `Update*/Build*/Create*()` | Звітні модулі | Кнопки «Оновити» на аркушах звітів |
| `HighlightRowsWithAllChecks()` | checkErrors.bas | Кнопка або Ribbon |

---

## 🔄 Спільні патерни коду

### Патерн 1: Пошуковий вибір (UserForm1 + clsSearchableDropdown)
```vba
Dim frm As UserForm1
Set frm = New UserForm1
frm.Caption = "Виберіть підрозділ"
frm.ListData = ws.Range("A1:A" & lastRow)  ' або Range або Array
frm.Show
If frm.ListBox1.ListIndex <> -1 Then
    selectedValue = frm.ListBox1.Value
Else
    MsgBox "Не вибрано!": Exit Sub
End If
```

### Патерн 2: Агрегація через Scripting.Dictionary
```vba
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
key = itemName & "||" & CStr(price) & "||" & uom
If Not dict.Exists(key) Then
    dict.Add key, 0#
End If
dict(key) = dict(key) + qty
```

### Патерн 3: Save-патерн (збереження документа)
```vba
' 1. Копіюємо аркуш
ThisWorkbook.Sheets("НазваАркуша").Copy
Dim WbNew As Workbook
Set WbNew = ActiveWorkbook
' 2. Видаляємо фігури/кнопки
For Each shp In WbNew.Worksheets(1).Shapes: shp.Delete: Next shp
' 3. Замінюємо формули на значення
For Each cell In WbNew.Worksheets(1).UsedRange
    If cell.HasFormula Then cell.Value = cell.Value
Next cell
' 4. Зберігаємо
Application.DisplayAlerts = False
WbNew.SaveAs FileName:=savePath, FileFormat:=xlOpenXMLWorkbook
Application.DisplayAlerts = True
WbNew.Close SaveChanges:=False
```

### Патерн 4: ProgressBar
```vba
ProgressBar.Progress percentage:=Int(r / totalRows * 100), _
    UnloadAfterCompletion:=(r = totalRows)
DoEvents
```

### Патерн 5: GetResidusePool (вибір пулу залишків)
```vba
Dim poolRange As Range
Set poolRange = GetResidusePool(wsResiduse2, FromUnitNamedRange.Value)
' Повертає видимі рядки залишків для підрозділу
' З питанням користувачу: "Лише свої залишки чи всі?"
```

---

## ⚠️ Відомі проблеми

| Пріоритет | Файл | Проблема |
|-----------|------|---------|
| 🔴 Критична | ResiduseOneGood.bas | Незадекларована змінна `selectedUnit` |
| 🔴 Критична | WOff81.bas, RequestM.bas | Незадекларована змінна `quantity` |
| 🔴 Критична | SaveInventoryNCA() | Відсутній `Application.DisplayAlerts = False` перед SaveAs — діалог XLM ламає збереження |
| 🟡 Серйозна | ManualAdd.bas | Кирилиця в іменах змінних |
| 🟡 Серйозна | WOff431.bas | `WbNew.Close` без перевірки чи WbNew існує |
| 🟡 Серйозна | RequestM.bas | `GoTo` у циклі — антипатерн |
| 🟢 Незначна | Reiduse2.bas, Add14.bas | Дублювання функцій SafeNum/SafeStr |
| 🟢 Незначна | UniversalImport.bas | `wsActualDoc.Activate` у циклі — уповільнення |
| 🟢 Незначна | checkErrors.bas | IsValidNumber — лише кома, ігнорує локаль |

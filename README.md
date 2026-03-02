# Navis3dExporter — плагин экспорта GLB для Navisworks 2020

Проект на C# для создания плагина Navisworks, который будет экспортировать геометрию текущей модели в формат GLB.

## Требования

- Autodesk Navisworks Manage/Simulate 2020 (x64)
- .NET Framework 4.7
- Visual Studio 2019/2022 (или совместимая)

## Сборка

1. Откройте `Navis3dExporter.csproj` в Visual Studio.
2. Проверьте, что ссылки на библиотеки Navisworks указывают на корректный каталог, например:
   - `C:\Program Files\Autodesk\Navisworks Manage 2020\Autodesk.Navisworks.Api.dll`
   - `C:\Program Files\Autodesk\Navisworks Manage 2020\Autodesk.Navisworks.Api.Com.dll`
3. Соберите проект в конфигурации `Release` и платформе `x64`.

## Размещение плагина

После сборки вы получите DLL `Navis3dExporter.dll`. Для загрузки плагина Navisworks также понадобится `.addin`‑файл (будет подготовлен на следующем шаге), который нужно поместить в одну из папок:

- `%APPDATA%\Autodesk\Navisworks Manage 2020\Plugins`
- или общесистемную папку плагинов Navisworks.

## Дальнейшие шаги

- Реализовать экспорт геометрии в GLB.
- Добавить `.addin`‑файл описания плагина.


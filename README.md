# Markdown to Word Converter

This repository contains a Python-based application that converts **Markdown** text into a **Word (.docx)** document using a user-friendly interface built with **Flet**.

You can use this to transfer Markdown tables from ChatGPT to Word.

## Features

- **Markdown to Word conversion**: Supports text, headings, and tables.
- **Customizable file saving**: Specify the output file name and directory.
- **Unique file naming**: Automatically handles duplicate file names by creating unique copies.
- **Beautiful Word formatting**: Custom styles for headings and table borders.

## Technologies Used

- **[Flet](https://flet.dev/)**: For building a modern GUI application.
- **[Markdown2](https://github.com/trentm/python-markdown2)**: For converting Markdown to HTML.
- **[BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/)**: For parsing and manipulating HTML.
- **[python-docx](https://python-docx.readthedocs.io/)**: For generating and styling Word documents.

## How It Works

1. Enter Markdown text in the input field.
2. Specify the desired file name and output directory.
3. Click **Convert** to create a Word document.
4. The app converts the Markdown to HTML, parses it, and generates a Word file with appropriate styles and formatting.

## Installation

### Prerequisites

- Python 3.9
- Install the required libraries using `pip`:

  ```bash
  pip install flet markdown2 beautifulsoup4 python-docx

### Clone the Repository
git clone https://github.com/miosion/markdown-to-word-converter.git

cd markdown-to-word-converter

### Usage
Run the application:

python MDtoWord.py

Interact with the app:

Paste your Markdown content into the input field.
Specify a file name and save path (optional).
Click Convert.
The Word document will be saved in the specified directory.

# Конвертер Markdown в Word

Данное приложение на Python позволяет конвертировать текст в формате **Markdown** в документ **Word (.docx)** через удобный интерфейс, созданный с помощью **Flet**.

Вы можете использовать его, чтобы перенести Markdown таблицы из ChatGPT в Word.

## Возможности

- **Конвертация Markdown в Word**: Поддерживает текст, заголовки и таблицы.
- **Настраиваемое сохранение файла**: Возможность задать имя файла и каталог для сохранения.
- **Уникальные имена файлов**: Автоматически создаёт уникальные копии при совпадении имён.
- **Красивая стилизация Word-файлов**: Пользовательские стили для заголовков и границ таблиц.

## Используемые технологии

- **[Flet](https://flet.dev/ru/)**: Для создания современного графического интерфейса.
- **[Markdown2](https://github.com/trentm/python-markdown2)**: Для конвертации Markdown в HTML.
- **[BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/)**: Для парсинга и обработки HTML.
- **[python-docx](https://python-docx.readthedocs.io/ru/latest/)**: Для генерации и стилизации документов Word.

## Как это работает

1. Введите текст в формате Markdown в текстовое поле.
2. Укажите желаемое имя файла и директорию для сохранения.
3. Нажмите **Конвертировать**, чтобы создать документ Word.
4. Приложение преобразует Markdown в HTML, парсит его и создаёт Word-файл с нужными стилями и форматированием.

## Установка

### Предварительные требования

- Python 3.9
- Установите необходимые библиотеки через `pip`:

  ```bash
  pip install flet markdown2 beautifulsoup4 python-docx

### Клонирование репозитория

git clone https://github.com/miosion/markdown-to-word-converter.git

cd markdown-to-word-converter

### Использование
Запустите приложение:

python MDtoWord.py
Работа с приложением:

Вставьте содержимое Markdown в текстовое поле.
Укажите имя файла и путь сохранения (по желанию).
Нажмите Конвертировать.
Документ Word будет сохранён в указанной директории.

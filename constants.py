ADVICE = "Реши задачу по математике."
"Ответ должен быть полным и пошаговым."
"Если текст задачи, несмотря на контекст непонятен, то таком случае верни текст 'Некорректное условие задачи'"
ANSWER_COLUMN = "answer"
AUTHOR = ' А. В. Шевкин.'
AUTHOR_SHEET_NAME = "author"
BATCH_SIZE = 32
CLASSES = '5;6'
CLASSES_COLUMN = "classes"
DESCRIPTION = 'Сборник включает текстовые задачи по разделам школьной математики: натуральные числа, дроби, пропорции, проценты, уравнения. ' \
'Ко многим задачам даны ответы или советы с чего начать решения. '
'Решения некоторых задач приведены в качестве образцов в основном тексте книги или в разделе «Ответы, советы, решения». '
'Материалы сборника можно использовать как дополнение к любому действующему учебнику. '
'При подготовке этого издания добавлены новые задачи и решения некоторых задач. '
'Пособие предназначено для учащихся 5–6 классов общеобразовательных школ, учителей, студентов педагогических вузов. '
DEST_FOLDER = "./artefacts_pytorch"
DOCX_PATH = "tekstovye_zadachi_po_matematike_1.docx"
GOOGLE_DRIVE_COMMON_PATH = "https://drive.google.com/uc?id"
ID_TASK_COLUMN = "id_tasks_book"
MISTRAL_MODEL = "mistral-large-latest"
NAME = 'Текстовые задачи по математике. 5–6 классы / А. В. Шевкин. — 3-е изд., перераб. — М. : Илекса, 2024. — 160 с. : ил.'
OUTPUT_FILE="tasks.xlsx"
PARAGRAPH_COLUMN = "paragraph"
SOLUTION_COLUMN = "AI_solution"
TASK_COLUMN = "task"
TASK_SHEET_NAME = "tasks"
TASK_SLICE_LENGTH = 50
TIME_SLEEP = 3
TOC_SHEET_NAME = "table_of_contents"
TOPIC_ID = 1
TRIM_CHARS = 5
AUTHOR_DATA = [
        {'name': NAME,
        'author': AUTHOR,
        'description': DESCRIPTION,
        'topic_id': TOPIC_ID,
        'classes': CLASSES
        }
    ]

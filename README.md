# Руководство по запуску шаблонов документов

В этом руководстве описаны типы шаблонов и необходимые параметры для их корректной генерации через API.

## Получение ID шаблонов

Перед использованием получите актуальные ID шаблонов через эндпоинт:

```
GET http://localhost:8000/officesvc/templates
```

## 1. Документы для групп

### Список группы (group_list_template.docx)
```
GET http://localhost:8000/officesvc/report?template_id=2&group_id=1
```
**Параметры:**
- `group_id` - ID группы

### Список группы в Excel (Список группы.xlsx)
```
GET http://localhost:8000/officesvc/report?template_id=14&group_id=1
```
**Параметры:**
- `group_id` - ID группы

## 2. Журналы оценок

### Журналы A4/A5 (Журнал A4.xlsx, Журнал A5.xlsx)
```
GET http://localhost:8000/officesvc/report?template_id=1&group_id=1&discipline_id=3&start_date=2025-04-01
```
или
```
GET http://localhost:8000/officesvc/report?template_id=7&group_id=1&discipline_id=3&start_date=2025-04-01
```
**Параметры:**
- `group_id` - ID группы
- `discipline_id` - ID дисциплины
- `start_date` - дата начала (опционально)

## 3. Документы для дисциплин

### БРС (БРС_Методы_и_стандарты_программирования_1_536_Б.docx)
```
GET http://localhost:8000/officesvc/report?template_id=8&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Вопросы к экзамену (Список вопросом для экзамена или зачета.docx)
```
GET http://localhost:8000/officesvc/report?template_id=23&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Список литературы (Список литературы.docx)
```
GET http://localhost:8000/officesvc/report?template_id=15&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Темы практических работ (Темы практических работ.docx)
```
GET http://localhost:8000/officesvc/report?template_id=22&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Шаблон билета (Шаблон билета.docx)
```
GET http://localhost:8000/officesvc/report?template_id=9&discipline_id=3&ticket_number=5
```
**Параметры:**
- `discipline_id` - ID дисциплины
- `ticket_number` - Номер билета (опционально)

## 4. Документы для студентов

### Отчет по практике (Отчет по практике 7 семестр.docx)
```
GET http://localhost:8000/officesvc/report?template_id=3&student_id=1&teacher_id=1&group_id=1
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `group_id` - ID группы (опционально)

### Отчет по практике (Отчет по практике 6 семестр.docx)
```
GET http://localhost:8000/officesvc/report?template_id=6&student_id=1&teacher_id=1&group_id=1
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `group_id` - ID группы (опционально)

### Отчет по технологической практике (Отчет по технологической практике.docx)
```
GET http://localhost:8000/officesvc/report?template_id=13&student_id=1&teacher_id=1&group_id=1
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `group_id` - ID группы (опционально)

### Титульные листы лабораторных работ
```
GET http://localhost:8000/officesvc/report?template_id=12&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Титульный лист Курсовая работа
```
GET http://localhost:8000/officesvc/report?template_id=5&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Титульный лист Курсовой проект
```
GET http://localhost:8000/officesvc/report?template_id=16&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Титульный лист Реферат
```
GET http://localhost:8000/officesvc/report?template_id=21&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Титульные листы для бакалавров и магистров
```
GET http://localhost:8000/officesvc/report?template_id=20&student_id=1&teacher_id=1&discipline_id=3
```
или
```
GET http://localhost:8000/officesvc/report?template_id=17&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Список публикаций (Список публикаций.docx)
```
GET http://localhost:8000/officesvc/report?template_id=19&student_id=1
```
**Параметры:**
- `student_id` - ID студента

## 5. Задание на выполнение выпускной работы

### Задание 4 курс, Задание 6 курс (ЗАДАНИЕ 4 курс.docx, ЗАДАНИЕ 6 курс.docx)
```
GET http://localhost:8000/officesvc/report?template_id=4&student_id=1&teacher_id=1&discipline_id=3
```
или
```
GET http://localhost:8000/officesvc/report?template_id=18&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины (опционально)

## 6. Расписания

### Расписание преподавателей (Расписание преподавателей.docx)
```
GET http://localhost:8000/officesvc/report?template_id=10&teacher_id=1&day_of_week=Понедельник
```
**Параметры:**
- `teacher_id` - ID преподавателя
- `day_of_week` - День недели (опционально)

### Загруженность аудиторий (Загруженность аудиторий.docx)
```
GET http://localhost:8000/officesvc/report?template_id=24&classroom_id=1&day_of_week=Вторник&group_id=1&group_id_2=2
```
**Параметры:**
- `classroom_id` - ID аудитории
- `day_of_week` - День недели (опционально)
- `group_id` - ID первой группы (опционально)
- `group_id_2` - ID второй группы (опционально)

## 7. Программы секций

### Программа секции ИТиПОРЭА (Программа секции ИТиПОРЭА.docx)
```
GET http://localhost:8000/officesvc/report?template_id=11
```
**Параметры:** не требуются

## Дополнительные параметры

Для любого запроса можно добавить параметр `download=true`, чтобы сразу скачать документ:
```
GET http://localhost:8000/officesvc/report?template_id=1&group_id=1&download=true
```

## Таблица соответствия шаблонов и параметров

| ID | Тип документа | Шаблон | Необходимые параметры |
|----|---------------|--------|------------------------|
| 1 | Журнал оценок A4 | Журнал A4.xlsx | group_id, discipline_id, start_date (опц.) |
| 2 | Список группы | group_list_template.docx | group_id |
| 3 | Отчет по практике 7 сем | Отчет по практике 7 семестр.docx | student_id, teacher_id, group_id (опц.) |
| 4 | Задание 4 курс | ЗАДАНИЕ 4 курс.docx | student_id, teacher_id, discipline_id (опц.) |
| 5 | Титульный лист курсовой | Титульный лист Курсовая работа.docx | student_id, teacher_id, discipline_id |
| 6 | Отчет по практике 6 сем | Отчет по практике 6 семестр.docx | student_id, teacher_id, group_id (опц.) |
| 7 | Журнал оценок A5 | Журнал A5.xlsx | group_id, discipline_id, start_date (опц.) |
| 8 | БРС | БРС_Методы_и_стандарты_программирования_1_536_Б.docx | discipline_id |
| 9 | Экзаменационный билет | Шаблон билета.docx | discipline_id, ticket_number (опц.) |
| 10 | Расписание преподавателей | Расписание преподавателей.docx | teacher_id, day_of_week (опц.) |
| 11 | Программа секции | Программа_секции_ИТиПОРЭА.docx | - |
| 12 | Титульный лист лабораторной | Титульный лист Лабораторная работа.docx | student_id, teacher_id, discipline_id |
| 13 | Отчет по технологической практике | Отчет по технологической практике.docx | student_id, teacher_id, group_id (опц.) |
| 14 | Список группы (Excel) | Список группы.xlsx | group_id |
| 15 | Список литературы | Список литературу.docx | discipline_id |
| 16 | Титульный лист курсового проекта | Титульный лист Курсовой проект.docx | student_id, teacher_id, discipline_id |
| 17 | Титульный лист для магистра | Шаблон титул маг.docx | student_id, teacher_id, discipline_id |
| 18 | Задание 6 курс | ЗАДАНИЕ 6 курс.docx | student_id, teacher_id, discipline_id (опц.) |
| 19 | Список публикаций | Список публикаций.docx | student_id |
| 20 | Титульный лист для бакалавра | Шаблон титул бал.docx | student_id, teacher_id, discipline_id |
| 21 | Титульный лист реферата | Титульный лист Реферат.docx | student_id, teacher_id, discipline_id |
| 22 | Темы практических работ | Темы практических работ.docx | discipline_id |
| 23 | Вопросы к экзамену | Список вопросом для экзамена или зачета.docx | discipline_id |
| 24 | Загруженность аудиторий | Загруженность аудиторий.docx | classroom_id, day_of_week (опц.), group_id (опц.), group_id_2 (опц.) |

## Примечания

1. Для получения ID групп, студентов, преподавателей, дисциплин и аудиторий используйте соответствующие эндпоинты:
   - `/officesvc/groups`
   - `/officesvc/students`
   - `/officesvc/teachers`
   - `/officesvc/disciplines`
   - `/officesvc/classrooms`
   - `/officesvc/days` - для получения списка дней недели
2. Вы можете скачать файл напрямую, добавив параметр `download=true` к любому запросу.
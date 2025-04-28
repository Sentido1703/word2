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
GET http://localhost:8000/officesvc/report?template_id=1&group_id=1
```
**Параметры:**
- `group_id` - ID группы

### Список группы в Excel (Список группы.xlsx)
```
GET http://localhost:8000/officesvc/report?template_id=8&group_id=1
```
**Параметры:**
- `group_id` - ID группы

## 2. Журналы оценок

### Журналы A4/A5 (Журнал A4.xlsx, Журнал A5.xlsx)
```
GET http://localhost:8000/officesvc/report?template_id=3&group_id=1&discipline_id=3&start_date=2025-04-01
```
**Параметры:**
- `group_id` - ID группы
- `discipline_id` - ID дисциплины
- `start_date` - дата начала (опционально)

## 3. Документы для дисциплин

### БРС (БРС_Методы_и_стандарты_программирования_1_536_Б.docx)
```
GET http://localhost:8000/officesvc/report?template_id=2&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Вопросы к экзамену (Список вопросом для экзамена или зачета.docx)
```
GET http://localhost:8000/officesvc/report?template_id=7&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Список литературы (Список литературу.docx)
```
GET http://localhost:8000/officesvc/report?template_id=9&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Темы практических работ (Темы практических работ.docx)
```
GET http://localhost:8000/officesvc/report?template_id=11&discipline_id=3
```
**Параметры:**
- `discipline_id` - ID дисциплины

### Шаблон билета (Шаблон билета.docx)
```
http://localhost:8000/officesvc/report?template_id=16&discipline_id=3&ticket_number=5
```
**Параметры:**
- `discipline_id` - ID дисциплины
- `ticket_number` - Номер билета

## 4. Документы для студентов

### Отчет по практике (Отчет по пратике 6/7 семестр.docx)
```
GET http://localhost:8000/officesvc/report?template_id=5&student_id=1&teacher_id=1&group_id=1
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `group_id` - ID группы (опционально)

### Титульные листы (курсовая, лабораторная, реферат)
```
GET http://localhost:8000/officesvc/report?template_id=12&student_id=1&teacher_id=1&discipline_id=3
```
**Параметры:**
- `student_id` - ID студента
- `teacher_id` - ID преподавателя
- `discipline_id` - ID дисциплины

### Список публикаций (Список публикаций.docx)
```
GET http://localhost:8000/officesvc/report?template_id=10&student_id=1
```
**Параметры:**
- `student_id` - ID студента

## Дополнительные параметры

Для любого запроса можно добавить параметр `download=true`, чтобы сразу скачать документ:
```
GET http://localhost:8000/officesvc/report?template_id=1&group_id=1&download=true
```

## Таблица соответствия шаблонов и параметров

| Тип документа | Шаблон | Необходимые параметры |
|---------------|--------|------------------------|
| Список группы | group_list_template.docx | group_id |
| Список группы (Excel) | Список группы.xlsx | group_id |
| Журнал оценок | Журнал A4.xlsx, Журнал A5.xlsx | group_id, discipline_id, start_date (опц.) |
| БРС | БРС_Методы_и_стандарты_программирования_1_536_Б.docx | discipline_id |
| Вопросы к экзамену | Список вопросом для экзамена или зачета.docx | discipline_id |
| Список литературы | Список литературу.docx | discipline_id |
| Темы практических работ | Темы практических работ.docx | discipline_id |
| Шаблон билета | Шаблон билета.docx | discipline_id |
| Отчет по практике | Отчет по пратике 6/7 семестр.docx | student_id, teacher_id, group_id |
| Титульный лист курсовой | Титульный лист Курсовая работа.docx | student_id, teacher_id, discipline_id |
| Титульный лист курсового проекта | Титульный лист Курсовой проект.docx | student_id, teacher_id, discipline_id |
| Титульный лист лабораторной | Титульный лист Лабороторная работа.docx | student_id, teacher_id, discipline_id |
| Титульный лист реферата | Титульный лист Реферат.docx | student_id, teacher_id, discipline_id |
| Список публикаций | Список публикаций.docx | student_id |

## Примечания

1. Номера `template_id` указаны условно. Используйте API `/officesvc/templates` для получения актуальных ID.
2. Для получения ID групп, студентов, преподавателей и дисциплин используйте соответствующие эндпоинты:
   - `/officesvc/groups`
   - `/officesvc/students`
   - `/officesvc/teachers`
   - `/officesvc/disciplines`
3. Вы можете скачать файл напрямую, добавив параметр `download=true` к любому запросу.
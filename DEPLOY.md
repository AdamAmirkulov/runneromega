# Деплой и обновление версий

Две машины:

| Машина    | Роль                    | Что делает                         |
|-----------|-------------------------|------------------------------------|
| Локальная | разработка              | правки кода → `git push`           |
| Сервер    | боевой запуск           | `git pull` (deploy.ps1) → рестарт  |

Единый источник правды — приватный репозиторий GitHub
`github.com/AdamAmirkulov/runneromega`.

---

## Что НЕ хранится в git (у каждой машины своё)

| Файл / папка                    | Почему                                   |
|---------------------------------|------------------------------------------|
| `scripts/config.py`             | пути и **пароли** — на сервере другие     |
| `scripts/service_account.json`  | ключ Google — секрет                      |
| `*.db`, `*.sqlite`              | живые данные (`users.db`, `my_database.sqlite`, `data/ais_oip.db`) |
| `data/runs/`, `uploads/`, `out/`| артефакты запусков                        |
| `.venv/`                        | окружение Python машины                   |
| `*.log`, `*.ipynb`, `*.bak`     | мусор / дев-файлы                         |

Шаблон конфига — `scripts/config.example.py` (в git). На новой машине:
`copy scripts\config.example.py scripts\config.py` и заполнить.

---

## Обычный цикл обновления

**Локально:**
```powershell
# поправил код…
.\push.ps1 "починил каскад суда в подаче иска"
```
(или вручную: `git add -A; git commit -m "…"; git push`)

**На сервере:**
```powershell
cd C:\путь\к\runneromega
.\deploy.ps1        # git pull + при необходимости pip install
# затем перезапустить приложение (см. ниже)
```

`deploy.ps1` не трогает `config.py`, базы и `data/` — они в `.gitignore`.

---

## Запуск приложения на сервере

Сейчас — вручную:
```powershell
.\run-server.ps1      # uvicorn app:app на 0.0.0.0:8000, без auto-reload
```

Чтобы крутилось само и рестартилось одной командой — можно завести
службу Windows (NSSM). Ставится один раз, тогда `deploy.ps1` сможет
делать `Restart-Service`. Скажи — настроим.

---

## Разовая миграция сервера (после чистки репозитория)

История репозитория переписана (из неё удалён `config.py` с паролями),
поэтому обычный `git pull` на сервере не пройдёт fast-forward. Один раз:

```powershell
cd C:\путь\к\runneromega

# 1. СОХРАНИТЬ локальные данные сервера в сторонку
mkdir ..\runneromega_backup
copy scripts\config.py              ..\runneromega_backup\
copy scripts\service_account.json   ..\runneromega_backup\ 2>$null
copy *.db                           ..\runneromega_backup\ 2>$null
copy *.sqlite                       ..\runneromega_backup\ 2>$null
copy data\ais_oip.db                ..\runneromega_backup\ 2>$null

# 2. Забрать переписанную историю
git fetch origin
git reset --hard origin/master
git clean -fdx -e .venv -e scripts/config.py -e "*.db" -e "*.sqlite" -e data/

# 3. Вернуть данные на место
copy ..\runneromega_backup\config.py            scripts\
copy ..\runneromega_backup\service_account.json scripts\ 2>$null
copy ..\runneromega_backup\*.db                 . 2>$null
copy ..\runneromega_backup\*.sqlite             . 2>$null
copy ..\runneromega_backup\ais_oip.db           data\ 2>$null

# 4. Проверить и запустить
.\run-server.ps1
```

Если на сервере папка проекта **без git вообще** — проще клонировать заново
рядом, перенести туда `config.py` / базы / `data/`, переключить запуск на
новую папку.

Дальше — только `deploy.ps1`.

---

## Безопасность (сделать сейчас)

1. **Сменить все пароли**, которые были в публичном `config.py`:
   `aisoip_password` и `sk_password` для всех компаний (1–4).
   Старые значения утекли в публичный git и кэши GitHub — считать
   скомпрометированными.
2. Репозиторий → **Private** (GitHub → Settings → General → Danger Zone →
   Change repository visibility).
3. Проверить, что новый `config.py` есть только локально и на сервере,
   и нигде не коммитится (`git status` должен его игнорировать).

# Деплой и обновление версий

Две машины:

| Машина    | Роль            | Что делает                        |
|-----------|-----------------|-----------------------------------|
| Локальная | разработка      | правки кода → `git push`          |
| Сервер    | боевой запуск   | `git pull` (deploy.ps1) → рестарт |

Единый источник правды — приватный репозиторий GitHub
`github.com/AdamAmirkulov/runneromega`, ветка `main`.

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
.\push.ps1 "починил каскад суда в подаче иска"
```
(или вручную: `git add -A; git commit -m "..."; git push origin main`)

**На сервере:**
```powershell
cd C:\Users\Администратор\Desktop\runneromega-main
.\deploy.ps1        # git pull + при необходимости pip install
# затем перезапустить приложение (см. ниже)
```

`deploy.ps1` не трогает `config.py`, базы и `data/` — они в `.gitignore`.

---

## Окружение Python

Изолированный `.venv` (Python 3.14). Все пакеты в `requirements.txt` идут
с готовыми колёсами под 3.14 — компилятор не нужен. Создать/пересоздать:
```powershell
powershell -ExecutionPolicy Bypass -File .\setup-venv.ps1
```
`.venv` в `.gitignore` — на каждой машине свой.

## Запуск приложения на сервере

Вручную:
```powershell
powershell -ExecutionPolicy Bypass -File .\run-server.ps1
```
(uvicorn app:app на 0.0.0.0:8000, без auto-reload; берёт `.venv`, иначе системный python)

Чтобы крутилось само и рестартилось одной командой — можно завести
службу Windows (NSSM). Тогда `deploy.ps1` сможет делать `Restart-Service`.

---

## Разовая миграция сервера (после чистки репозитория)

История переписана (удалён `config.py` с паролями, ветка `master` → `main`),
поэтому обычный `git pull` не проходит fast-forward. Разовый скрипт делает
всё сам:

```powershell
cd C:\Users\Администратор\Desktop\runneromega-main
# закрыть запущенный uvicorn
git fetch origin
git checkout origin/main -- migrate-server-once.ps1
powershell -ExecutionPolicy Bypass -File .\migrate-server-once.ps1
```

Он: бэкапит `config.py` + базы + ключ в `..\_runneromega_backup`,
переключает git на чистую `main`, чистит мусор, возвращает данные на место,
пересоздаёт `.venv` и ставит зависимости. `data/`, `uploads/`, `logs/` и
сами базы не трогаются.

Дальше — только `deploy.ps1`.

---

## Безопасность

1. **Сменить пароли**, которые были в публичном `config.py`:
   `aisoip_password` и `sk_password` для всех компаний (1–4).
   Старые значения были в публичном git — считать скомпрометированными.
2. Репозиторий → **Private** (GitHub → Settings → General → Danger Zone →
   Change repository visibility).
3. `git status` не должен показывать `scripts/config.py` — он в `.gitignore`.

# SEEA Daily Monitor

Ежедневно изтегля публичната таблица от:

`https://portal.seea.government.bg/bg/ByProducerAndEnergyObject`

Скриптът:

- зарежда актуалните записи;
- създава стабилен `record_id` за всеки ред;
- добавя само новите записи;
- обновява `last_seen_at` за записите, които все още присъстват;
- записва резултатите в CSV и Excel;
- commit-ва промените автоматично чрез GitHub Actions.

## Файлове с данни

- `data/energy_objects.csv`
- `data/energy_objects.xlsx`
- `data/last_run.json`

## Първоначално стартиране

1. Създай ново **private** GitHub repository.
2. Качи всички файлове от тази папка.
3. Отвори `Settings → Actions → General`.
4. В `Workflow permissions` избери **Read and write permissions**, ако настройките на repository-то не позволяват `contents: write`.
5. Отвори `Actions → Daily SEEA update → Run workflow`.
6. След успешното изпълнение данните ще се появят в папка `data/`.

Не са необходими GitHub Secrets, защото порталът е публичен, а workflow-ът използва вградения `GITHUB_TOKEN`.

## Локален тест

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scraper.py
```

При Windows активирай средата с:

```powershell
.venv\Scripts\Activate.ps1
```

## Важно

Ако порталът промени HTML таблицата или endpoint-а, workflow-ът ще завърши с грешка, вместо да презапише файла с празни данни.

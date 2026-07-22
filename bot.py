#!/usr/bin/env python3
import os
import sys
import argparse
from dotenv import load_dotenv
from yandex_shared_disk import YandexSharedDiskClient, YandexSharedDiskError

def init_client(args: argparse.Namespace) -> YandexSharedDiskClient:
    # Загружаем переменные из файла .env, если он существует
    load_dotenv()

    token = args.token or os.getenv("YANDEX_OAUTH_TOKEN")
    vd_hash = args.vd_hash or os.getenv("YANDEX_VD_HASH")

    if not token:
        print(" Ошибка: OAuth-токен Яндекс Диска не найден.")
        print("Укажите YANDEX_OAUTH_TOKEN в файле .env или передайте через аргумент --token.")
        sys.exit(1)

    if not vd_hash:
        print(" Ошибка: Метка общего диска (vd_hash) не найдена.")
        print("Укажите YANDEX_VD_HASH в файле .env или передайте через аргумент --vd-hash.")
        sys.exit(1)

    return YandexSharedDiskClient(oauth_token=token, vd_hash=vd_hash)

def main():
    parser = argparse.ArgumentParser(
        description="Бот/CLI для скачивания файлов с общего Яндекс Диска (Яндекс 360 для бизнеса)."
    )
    
    # Общие аргументы аутентификации
    parser.add_argument("--token", help="OAuth-токен Яндекс Диска (если не задан в .env)")
    parser.add_argument("--vd-hash", help="Метка общего диска vd_hash (если не задана в .env)")

    subparsers = parser.add_subparsers(dest="command", help="Команда для выполнения")

    # Команда 1: download
    download_parser = subparsers.add_parser("download", help="Скачать один файл с общего диска")
    download_parser.add_argument(
        "remote_path",
        help="Путь к файлу на общего диске (напр. 'folder/file.pdf' или 'test.txt')"
    )
    download_parser.add_argument(
        "-o", "--output",
        help="Локальный путь для сохранения файла (по умолчанию - имя исходного файла)"
    )

    # Команда 2: list
    list_parser = subparsers.add_parser("list", help="Просмотреть содержимое папки на общем диске")
    list_parser.add_argument(
        "remote_folder",
        nargs="?",
        default="/",
        help="Путь к папке на общем диске (по умолчанию '/')"
    )

    # Команда 3: download-folder
    df_parser = subparsers.add_parser("download-folder", help="Скачать содержимое всей папки с общего диска")
    df_parser.add_argument(
        "remote_folder",
        help="Путь к папке на общем диске для скачивания"
    )
    df_parser.add_argument(
        "-o", "--output-dir",
        default="./downloads",
        help="Локальная директория для сохранения (по умолчанию './downloads')"
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    try:
        client = init_client(args)

        if args.command == "download":
            print(f" Начинаем скачивание файла: {args.remote_path}")
            saved_path = client.download_file(
                remote_path=args.remote_path,
                save_path=args.output
            )
            print(f" Успешно скачано: {saved_path}")

        elif args.command == "list":
            print(f" Получение списка элементов в '{args.remote_folder}'...")
            items = client.list_folder(args.remote_folder)
            if not items:
                print("Папка пуста или не содержит файлов.")
            else:
                print(f"\n{'ТИП':<6} | {'РАЗМЕР':<10} | {'ИМЯ'}")
                print("-" * 50)
                for item in items:
                    t = item.get("type", "file")
                    size = f"{item.get('size', 0):,} B" if t == "file" else "-"
                    name = item.get("name")
                    print(f"{t:<6} | {size:<10} | {name}")
                print(f"\nВсего элементов: {len(items)}")

        elif args.command == "download-folder":
            print(f" Начинаем скачивание папки '{args.remote_folder}' в '{args.output_dir}'...")
            downloaded = client.download_folder(
                remote_folder=args.remote_folder,
                local_dest=args.output_dir
            )
            print(f" Скачивание завершено! Скачано файлов: {len(downloaded)}")
            for f in downloaded:
                print(f"  - {f}")

    except YandexSharedDiskError as e:
        print(f" Ошибка Яндекс Диска: {e}")
        sys.exit(1)
    except Exception as e:
        print(f" Произошла непредвиденная ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

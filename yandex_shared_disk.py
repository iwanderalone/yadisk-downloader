import os
import sys
import logging
from typing import Optional, List, Dict, Any
import requests
from tqdm import tqdm

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("YandexSharedDisk")

API_BASE_URL = "https://cloud-api.yandex.net/v1/disk/virtual-disks"

class YandexSharedDiskError(Exception):
    """Базовое исключение для ошибок Яндекс 360 / Общего диска."""
    pass

class YandexSharedDiskClient:
    """
    Клиент для работы с Общими дисками Яндекс 360 для бизнеса (Virtual Disks API).
    Документация: https://yandex.ru/dev/disk-api/doc/ru/reference/content_shd
    """
    def __init__(self, oauth_token: str, vd_hash: str):
        if not oauth_token:
            raise ValueError("OAuth-токен не передан")
        if not vd_hash:
            raise ValueError("Метка общего диска (vd_hash) не передана")
        
        self.oauth_token = oauth_token.strip()
        self.vd_hash = vd_hash.strip()
        
        self.headers = {
            "Authorization": f"OAuth {self.oauth_token}",
            "Accept": "application/json"
        }

    def format_vd_path(self, remote_path: str) -> str:
        """
        Форматирует локальный путь файла на общем диске к виду:
        vd:<vd_hash>:disk:/<путь внутри общего диска>
        """
        remote_path = remote_path.strip()
        if remote_path.startswith("vd:"):
            return remote_path
        
        # Убираем начальный слэш если есть, чтобы корректно собрать путь
        clean_path = remote_path.lstrip("/")
        return f"vd:{self.vd_hash}:disk:/{clean_path}"

    def get_download_url(self, remote_path: str) -> str:
        """
        Запрашивает URL для скачивания файла с общего диска.
        Endpoint: GET https://cloud-api.yandex.net/v1/disk/virtual-disks/resources/download
        """
        formatted_path = self.format_vd_path(remote_path)
        url = f"{API_BASE_URL}/resources/download"
        params = {"path": formatted_path}

        logger.info("Запрос ссылки на скачивание для: %s", formatted_path)
        response = requests.get(url, headers=self.headers, params=params, timeout=30)

        if response.status_code == 200:
            data = response.json()
            href = data.get("href")
            if not href:
                raise YandexSharedDiskError(f"Ответ сервера не содержит 'href': {data}")
            return href
        elif response.status_code == 401:
            raise YandexSharedDiskError("401 Unauthorized: Проверьте правильность OAuth-токена.")
        elif response.status_code == 403:
            raise YandexSharedDiskError("403 Forbidden: Нет доступа к данному общему диску или ресурсу.")
        elif response.status_code == 404:
            raise YandexSharedDiskError(f"404 Not Found: Файл по пути '{formatted_path}' не найден.")
        else:
            try:
                err_data = response.json()
                msg = err_data.get("message", response.text)
            except Exception:
                msg = response.text
            raise YandexSharedDiskError(f"Ошибка API ({response.status_code}): {msg}")

    def download_file(
        self,
        remote_path: str,
        save_path: Optional[str] = None,
        show_progress: bool = True
    ) -> str:
        """
        Скачивает файл с общего диска по указанному пути.
        
        :param remote_path: Путь к файлу на общем диске (напр. 'folder/file.pdf' или 'file.txt')
        :param save_path: Локальный путь сохранения. Если не указан, берется имя файла из remote_path.
        :param show_progress: Отображать ли прогресс-бар скачивания.
        :return: Итоговый путь к сохраненному локальному файлу.
        """
        download_url = self.get_download_url(remote_path)

        if not save_path:
            save_path = os.path.basename(remote_path.rstrip("/"))
            if not save_path:
                save_path = "downloaded_file"

        # Создаем необходимые директории для локального пути
        output_dir = os.path.dirname(os.path.abspath(save_path))
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        logger.info("Скачивание файла в '%s'...", save_path)
        
        # Согласно документации, при запросе скачивания следует передавать тот же OAuth-токен
        download_headers = {
            "Authorization": f"OAuth {self.oauth_token}"
        }

        with requests.get(download_url, headers=download_headers, stream=True, timeout=60) as resp:
            resp.raise_for_status()
            total_size = int(resp.headers.get("content-length", 0))

            chunk_size = 1024 * 1024  # 1MB
            
            if show_progress:
                pbar = tqdm(
                    total=total_size if total_size > 0 else None,
                    unit="B",
                    unit_scale=True,
                    desc=os.path.basename(save_path)
                )
            else:
                pbar = None

            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=chunk_size):
                    if chunk:
                        f.write(chunk)
                        if pbar:
                            pbar.update(len(chunk))
            
            if pbar:
                pbar.close()

        logger.info("Файл успешно сохранен: %s", save_path)
        return save_path

    def get_resource_info(self, remote_path: str = "/") -> Dict[str, Any]:
        """
        Получает метаинформацию о файле или папке на общем диске.
        Endpoint: GET https://cloud-api.yandex.net/v1/disk/virtual-disks/resources
        """
        formatted_path = self.format_vd_path(remote_path)
        url = f"{API_BASE_URL}/resources"
        params = {"path": formatted_path}

        response = requests.get(url, headers=self.headers, params=params, timeout=30)
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 404:
            raise YandexSharedDiskError(f"Ресурс '{formatted_path}' не найден (404).")
        else:
            raise YandexSharedDiskError(f"Ошибка получения инфо ({response.status_code}): {response.text}")

    def list_folder(self, remote_folder: str = "/") -> List[Dict[str, Any]]:
        """
        Возвращает список предметов (файлов и подпапок) в указанной папке на общем диске.
        """
        info = self.get_resource_info(remote_folder)
        _embedded = info.get("_embedded", {})
        items = _embedded.get("items", [])
        return items

    def download_folder(self, remote_folder: str, local_dest: str = "./downloads") -> List[str]:
        """
        Рекурсивно или пакетом скачивает все файлы из указанной папки на общего диске.
        """
        items = self.list_folder(remote_folder)
        downloaded_files = []

        for item in items:
            name = item.get("name")
            item_type = item.get("type")
            item_path = item.get("path")
            
            # Извлекаем чистый путь относительно общего диска
            # path в ответах Яндекс API имеет вид: vd:<vd_hash>:disk:/folder/subfolder/file.ext
            rel_path = item_path.split(":disk:", 1)[-1] if ":disk:" in item_path else name
            
            local_target = os.path.join(local_dest, rel_path.lstrip("/"))

            if item_type == "file":
                saved = self.download_file(remote_path=item_path, save_path=local_target)
                downloaded_files.append(saved)
            elif item_type == "dir":
                sub_files = self.download_folder(remote_folder=item_path, local_dest=local_dest)
                downloaded_files.extend(sub_files)

        return downloaded_files

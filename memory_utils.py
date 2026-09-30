"""Recursos temporales para procesar registros sin retenerlos todos en RAM."""

from functools import wraps
from itertools import islice
from concurrent.futures import ThreadPoolExecutor
import pickle
from tempfile import TemporaryDirectory, TemporaryFile
from threading import BoundedSemaphore, local

import xlsxwriter


# Un proceso Gunicorn (configuración del proyecto), una tarea pesada a la vez.
# Los Excel en segundo plano esperan sin crear un hilo por cada solicitud.
export_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="excel")
_heavy_slot = BoundedSemaphore(1)
_active = local()


def memory_limited(wait=False):
    def decorate(function):
        @wraps(function)
        def run(*args, **kwargs):
            if getattr(_active, "inside", False):
                return function(*args, **kwargs)
            if not _heavy_slot.acquire(blocking=wait):
                from flask import Response
                return Response(
                    "Hay una importación o exportación en curso. Espera a que termine "
                    "y vuelve a intentarlo. Esta solicitud no ha modificado los registros.",
                    status=503, content_type="text/plain; charset=utf-8",
                    headers={"Retry-After": "30", "Cache-Control": "no-store"},
                )
            _active.inside = True
            try:
                return function(*args, **kwargs)
            finally:
                _active.inside = False
                _heavy_slot.release()
        return run
    return decorate


def send_disk_file(file, **kwargs):
    """Cierra el temporal al acabar la descarga o al desconectarse el cliente."""
    from flask import send_file

    class ClosingDownload:
        def __init__(self, iterable):
            self.iterable = iterable
            self.iterator = iter(iterable)

        def __iter__(self):
            return self

        def __next__(self):
            try:
                return next(self.iterator)
            except BaseException:
                self.close()
                raise

        def close(self):
            try:
                close = getattr(self.iterable, "close", None)
                if close:
                    close()
            finally:
                file.close()

    try:
        response = send_file(file, **kwargs)
        response.response = ClosingDownload(response.response)
        return response
    except BaseException:
        file.close()
        raise


class DiskRows:
    """Secuencia reiterable de filas internas; nunca deserializa archivos subidos.

    El archivo anónimo se elimina al cerrar su contexto, también ante errores.
    Conserva date/time y los tipos originales necesarios para exportar Excel.
    """

    def __init__(self):
        self.file = TemporaryFile(mode="w+b")
        self.count = 0

    def append(self, row):
        self.file.seek(0, 2)
        pickle.dump(row, self.file, protocol=pickle.HIGHEST_PROTOCOL)
        self.count += 1

    def __len__(self):
        return self.count

    def __iter__(self):
        self.file.seek(0)
        for _ in range(self.count):
            yield pickle.load(self.file)

    def __getitem__(self, key):
        if isinstance(key, slice):
            start, stop, step = key.indices(self.count)
            if step < 1:
                raise ValueError("Solo se admiten cortes de avance.")
            return list(islice(self, start, stop, step))
        if key < 0:
            key += self.count
        if key < 0 or key >= self.count:
            raise IndexError(key)
        return next(islice(self, key, key + 1))

    def close(self):
        self.file.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


def disk_workbook(writer):
    """Excel por filas y salida en disco; send_file cierra el archivo al finalizar."""
    @wraps(writer)
    def export(report):
        output = TemporaryFile(mode="w+b")
        try:
            with TemporaryDirectory(prefix="reportabilidad-excel-") as directory:
                with xlsxwriter.Workbook(output, {
                    "constant_memory": True, "tmpdir": directory,
                    "strings_to_formulas": False, "strings_to_urls": False,
                }) as book:
                    writer(report, book)
            output.seek(0)
            return output
        except BaseException:
            output.close()
            raise
    return export

from dataclasses import dataclass

import kiotviet


@dataclass
class RunStats:
    ok: int = 0
    skipped: int = 0
    error: int = 0
    stopped: bool = False


def run_items(driver, items, result_timeout, stop_event, log, progress, process=None):
    """Xử lý lần lượt từng mã hàng.

    log(message, level) với level: info | ok | warn | error.
    progress(done, total) được gọi sau mỗi mã.
    """
    process = process or kiotviet.process_item
    stats = RunStats()
    total = len(items)
    progress(0, total)

    for index, item in enumerate(items, start=1):
        if stop_event.is_set():
            stats.stopped = True
            break

        code = item["code"]
        label = f"Mã {code} (dòng {item['row']})"

        try:
            if item["qty"] is None:
                raise kiotviet.SkipItem("số lượng trong Excel không phải là số.")
            process(driver, code, item["qty"], result_timeout)
            stats.ok += 1
        except kiotviet.SkipItem as exc:
            stats.skipped += 1
            log(f"BỎ QUA: {label} - {exc}", "warn")
        except kiotviet.VerifyError as exc:
            stats.error += 1
            log(f"CẦN KIỂM TRA: {label} - {exc}", "error")
        except Exception as exc:
            stats.error += 1
            first_line = str(exc).split("\n")[0]
            log(f"LỖI Web: {label}. Chi tiết: {first_line or type(exc).__name__}", "error")
        finally:
            progress(index, total)

    return stats

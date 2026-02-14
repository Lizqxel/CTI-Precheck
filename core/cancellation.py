from services.area_search import clear_cancel_flag as clear_cancel_flag_west
from services.area_search import set_cancel_flag as set_cancel_flag_west
from services import area_search_east


def request_cancel_service() -> None:
    try:
        set_cancel_flag_west(True)
    except Exception:
        pass

    try:
        area_search_east.set_cancel_flag(True)
    except Exception:
        pass

    for module in (None, area_search_east):
        try:
            driver = None
            if module is None:
                from services import area_search

                driver = getattr(area_search, "global_driver", None)
            else:
                driver = getattr(module, "global_driver", None)

            if driver:
                driver.quit()
        except Exception:
            pass


def clear_cancel_flags() -> None:
    try:
        clear_cancel_flag_west()
    except Exception:
        pass

    try:
        area_search_east.clear_cancel_flag()
    except Exception:
        pass

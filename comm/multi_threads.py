import threading
import concurrent.futures


def multi_threads_request(thread_list: list):
    threads_list = []
    for value in thread_list:
        t = threading.Thread(target=value[0], args=value[1])
        threads_list.append(t)
        t.start()
    for thread in threads_list:
        thread.join()


def concurrent_futures_requestion(fn, args: list):
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fn, arg) for arg in args]
    concurrent.futures.wait(futures)

import json
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

from server import Store, make_server


class CounterTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / 'stats.sqlite3'
        self.store = Store(self.path)

    def tearDown(self):
        self.directory.cleanup()

    def test_concurrent_clicks_repeat_and_persist(self):
        with ThreadPoolExecutor(max_workers=16) as pool:
            list(pool.map(self.store.record, ['mac', 'windows'] * 50))
        data = Store(self.path).summary()
        self.assertEqual((data['total'], data['mac'], data['windows'], data['today']), (100, 50, 50, 100))
        self.assertEqual(data['total'], sum(d['total'] for d in data['daily']))
        self.assertEqual(data['started_at'], self.store.summary()['started_at'])

    def test_shanghai_midnight_and_history_windows(self):
        self.store.record('mac', datetime(2026, 9, 27, 15, 59, tzinfo=timezone.utc))
        self.store.record('windows', datetime(2026, 9, 27, 16, 0, tzinfo=timezone.utc))
        self.store.record('mac', datetime(2026, 9, 1, tzinfo=timezone.utc))
        self.store.record('mac', datetime(2026, 8, 1, tzinfo=timezone.utc))
        data = self.store.summary(datetime(2026, 9, 27, 17, tzinfo=timezone.utc))
        self.assertEqual(data['daily'][0], {'day': '2026-09-28', 'mac': 0, 'windows': 1, 'total': 1})
        self.assertEqual((data['total'], data['today'], data['last_7_days'], data['last_30_days']), (4, 1, 2, 3))

    def test_static_snapshot_preserves_counts_and_rolls_over_without_new_clicks(self):
        snapshot = Path(self.directory.name) / 'public' / 'data.json'
        before = datetime(2026, 9, 27, 15, 59, tzinfo=timezone.utc)
        after = datetime(2026, 9, 27, 16, 1, tzinfo=timezone.utc)
        self.store.record('mac', before)
        self.store.publish(snapshot, before)
        first = json.loads(snapshot.read_text())
        self.assertEqual((first['total'], first['today']), (1, 1))
        self.store.publish(snapshot, after)
        second = json.loads(snapshot.read_text())
        self.assertEqual((second['total'], second['today']), (1, 0))
        self.assertEqual(second['started_at'], first['started_at'])
        self.assertEqual(snapshot.stat().st_mode & 0o777, 0o644)
        self.assertEqual(list(snapshot.parent.iterdir()), [snapshot])

    def test_http_only_accepts_small_same_origin_platform_events(self):
        server = make_server(self.path, 0)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        base = 'http://127.0.0.1:' + str(server.server_port)

        def request(path, body=None, origin='https://thewriter.cn'):
            req = urllib.request.Request(base + path, data=body, headers={'Origin': origin, 'Content-Type': 'text/plain'})
            try:
                with urllib.request.urlopen(req) as response:
                    return response.status, response.read()
            except urllib.error.HTTPError as error:
                return error.code, error.read()

        try:
            self.assertEqual(request('/api/download-clicks', b'{"platform":"mac"}')[0], 204)
            self.assertEqual(request('/api/download-clicks', b'{"platform":"windows"}', 'https://www.thewriter.cn')[0], 204)
            for body in [b'[]', b'null', b'bad json', b'{"platform":"linux"}', b'{"platform":[]}', b'{"platform":"mac","id":"visitor"}']:
                self.assertEqual(request('/api/download-clicks', body)[0], 400)
            self.assertEqual(request('/api/download-clicks', b'X' * 129)[0], 413)
            self.assertEqual(request('/api/download-clicks', b'{"platform":"mac"}', 'https://example.org')[0], 403)
            self.assertEqual(request('/api/download-clicks')[0], 404)
            self.assertEqual(request('/api/download-stats')[0], 404)
            self.assertEqual(self.store.summary()['total'], 2)
        finally:
            server.shutdown()
            server.server_close()
            thread.join()


if __name__ == '__main__':
    unittest.main()

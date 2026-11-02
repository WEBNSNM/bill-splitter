import unittest
import hashlib
import tempfile
from pathlib import Path
from unittest.mock import patch
from io import BytesIO
from bill_splitter.updater import version_tuple, select_release
from bill_splitter.updater import verify_checksum, download_update, EXE_NAME, CHECKSUM_NAME


class UpdateTests(unittest.TestCase):
    def release(self):
        return {'tag_name': 'v9.0.0', 'assets': [
            {'name': name, 'browser_download_url': 'https://github.com/WEBNSNM/bill-splitter/releases/download/v9.0.0/' + name}
            for name in (EXE_NAME, CHECKSUM_NAME)]}

    def test_release_and_invalid_host(self):
        release = self.release()
        self.assertEqual(select_release(release)['version'], 'v9.0.0')
        release['prerelease'] = True
        self.assertIsNone(select_release(release))
        release['prerelease'] = False
        release['assets'][0]['browser_download_url'] = 'https://other.example/file.exe'
        with self.assertRaises(ValueError):
            select_release(release)

    def test_download_and_checksum(self):
        data = b'test-executable'
        digest = hashlib.sha256(data).hexdigest()
        with patch('bill_splitter.updater.fetch', side_effect=[BytesIO(digest.encode()), BytesIO(data)]):
            path = download_update(select_release(self.release()))
        self.addCleanup(lambda: __import__('shutil').rmtree(path.parent))
        self.assertEqual(path.read_bytes(), data)
        self.assertEqual(path.name, EXE_NAME)
        with self.assertRaises(ValueError):
            verify_checksum(path, '0' * 64)
        with self.assertRaises(ValueError):
            verify_checksum(path, '')

    def test_versions(self):
        self.assertGreater(version_tuple('v0.10.0'), version_tuple('0.9.0'))
        with self.assertRaises(ValueError):
            version_tuple('1.0-beta')

    def test_release_assets(self):
        release = {'tag_name': 'v0.3.0', 'assets': []}
        self.assertIsNone(select_release(release, '0.3.0'))
        with self.assertRaises(ValueError):
            select_release(release, '0.2.0')

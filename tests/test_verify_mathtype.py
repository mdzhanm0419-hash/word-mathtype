"""使用虚构 OLE 和 DOCX 样本检查对象结构校验，不依赖 Word 或 MathType。"""
import importlib.util
import json
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts' / 'verify-mathtype.py'
spec = importlib.util.spec_from_file_location('verify_mathtype', SCRIPT)
verify = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify)
FREE = 0xFFFFFFFF
END = 0xFFFFFFFE
FAT_SECTOR = 0xFFFFFFFD


def entry(name, kind, first, length):
    data = bytearray(128)
    encoded = (name + '\0').encode('utf-16le')
    data[:len(encoded)] = encoded
    struct.pack_into('<H', data, 64, len(encoded))
    data[66] = kind
    struct.pack_into('<I', data, 116, first)
    struct.pack_into('<Q', data, 120, length)
    return data


def fake_ole(*, declared=64, empty=False, mini_cycle=False, large=False):
    header = bytearray(512)
    header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<H', header, 26, 3)
    struct.pack_into('<H', header, 28, 0xFFFE)
    struct.pack_into('<H', header, 30, 9)
    struct.pack_into('<H', header, 32, 6)
    struct.pack_into('<I', header, 44, 1)
    struct.pack_into('<I', header, 48, 1)
    struct.pack_into('<I', header, 56, 4096)
    struct.pack_into('<I', header, 60, END if large else 2)
    struct.pack_into('<I', header, 64, 0 if large else 1)
    struct.pack_into('<I', header, 68, END)
    struct.pack_into('<109I', header, 76, 0, *([FREE] * 108))
    fat = [FREE] * 128
    fat[0] = FAT_SECTOR
    fat[1] = END
    directory = bytearray(512)
    if large:
        length = 4096
        directory[:128] = entry('Root Entry', 5, END, 0)
        directory[128:256] = entry('Equation Native', 2, 2, length)
        for n in range(2, 10):
            fat[n] = END if n == 9 else n + 1
        payload = bytearray(length)
        payload[30] = 1
        return bytes(header) + struct.pack('<128I', *fat) + bytes(directory) + bytes(payload)
    fat[2] = END
    fat[3] = END
    directory[:128] = entry('Root Entry', 5, 3, 512)
    directory[128:256] = entry('Equation Native', 2, 0, declared)
    minifat = [FREE] * 128
    minifat[0] = 0 if mini_cycle else END
    mini_stream = bytearray(512)
    if not empty:
        mini_stream[30] = 1
    return bytes(header) + struct.pack('<128I', *fat) + bytes(directory) + struct.pack('<128I', *minifat) + bytes(mini_stream)


def fake_docx(path, payload, progid='Equation.DSMT4'):
    document = (f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'xmlns:o="urn:schemas-microsoft-com:office:office" '
                f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<w:body><o:OLEObject ProgID="{progid}" r:id="rId1"/></w:body></w:document>')
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="embeddings/oleObject1.bin" Type="urn:test"/>'
            '</Relationships>')
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('word/document.xml', document)
        archive.writestr('word/_rels/document.xml.rels', rels)
        archive.writestr('word/embeddings/oleObject1.bin', payload)


class VerifyTests(unittest.TestCase):
    def test_reads_real_bytes_of_ministream(self):
        self.assertEqual(verify.native_stream_size(fake_ole()), 64)

    def test_rejects_declared_stream_longer_than_chain(self):
        with self.assertRaises(ValueError):
            verify.native_stream_size(fake_ole(declared=128))

    def test_rejects_missing_or_zero_content(self):
        for payload in (fake_ole(empty=True), fake_ole(declared=28), b'not an ole'):
            with self.subTest(length=len(payload)), self.assertRaises(ValueError):
                verify.native_stream_size(payload)

    def test_rejects_loop_and_truncated_sector(self):
        for payload in (fake_ole(declared=128, mini_cycle=True), fake_ole()[:-1]):
            with self.subTest(length=len(payload)), self.assertRaises(ValueError):
                verify.native_stream_size(payload)

    def test_reads_regular_stream(self):
        self.assertEqual(verify.native_stream_size(fake_ole(large=True)), 4096)

    def test_cli_detects_count_and_payload(self):
        with tempfile.TemporaryDirectory() as tmp:
            doc = Path(tmp) / 'synthetic.docx'
            fake_docx(doc, fake_ole())
            result = subprocess.run([sys.executable, str(SCRIPT), str(doc), '--expected-total', '1'], capture_output=True, text=True, encoding='utf-8')
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)['equation_native_sizes'], [64])
            result = subprocess.run([sys.executable, str(SCRIPT), str(doc), '--expected-total', '2'], capture_output=True)
            self.assertNotEqual(result.returncode, 0)


if __name__ == '__main__':
    unittest.main()

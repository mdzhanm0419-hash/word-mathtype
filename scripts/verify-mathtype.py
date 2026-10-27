"""检查指定占位符生成的 MathType OLE 数量及非空内部数据；不判断公式语义。"""
import argparse
import json
import posixpath
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

OLE = '{urn:schemas-microsoft-com:office:office}OLEObject'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
END = 0xFFFFFFFE
FREE = 0xFFFFFFFF


def native_stream_size(data):
    if data[:8] != bytes.fromhex('d0cf11e0a1b11ae1') or len(data) < 512:
        raise ValueError('嵌入对象不是有效的复合文件')
    shift = struct.unpack_from('<H', data, 30)[0]
    if shift not in (9, 12):
        raise ValueError('不支持的复合文件扇区大小')
    size = 1 << shift
    mini_shift = struct.unpack_from('<H', data, 32)[0]
    if mini_shift != 6:
        raise ValueError('不支持的迷你扇区大小')
    fat_count = struct.unpack_from('<I', data, 44)[0]
    first_dir = struct.unpack_from('<I', data, 48)[0]
    mini_cutoff = struct.unpack_from('<I', data, 56)[0]
    first_minifat = struct.unpack_from('<I', data, 60)[0]
    minifat_count = struct.unpack_from('<I', data, 64)[0]
    extra_difat_count = struct.unpack_from('<I', data, 72)[0]
    if extra_difat_count:
        raise ValueError('复合文件需要扩展 DIFAT，当前检查器不支持')
    sectors = [v for v in struct.unpack_from('<109I', data, 76) if v != FREE][:fat_count]
    if len(sectors) != fat_count:
        raise ValueError('FAT 扇区不完整')
    def sector(n):
        offset = (n + 1) * size
        part = data[offset:offset + size]
        if len(part) != size:
            raise ValueError('复合文件扇区越界')
        return part
    fat = []
    for n in sectors:
        fat.extend(struct.unpack('<' + 'I' * (size // 4), sector(n)))
    def chain(first, allocation, read_block, needed=None):
        data_out = bytearray()
        visited = set()
        n = first
        while n != END:
            if n in visited or n >= len(allocation):
                raise ValueError('复合文件扇区链异常')
            visited.add(n)
            data_out.extend(read_block(n))
            if needed is not None and len(data_out) >= needed:
                break
            n = allocation[n]
        if needed is not None and len(data_out) < needed:
            raise ValueError('复合文件流数据不完整')
        return bytes(data_out[:needed] if needed is not None else data_out)

    directory = chain(first_dir, fat, sector)
    found = []
    root = []
    for offset in range(0, len(directory), 128):
        entry = directory[offset:offset + 128]
        name_len = struct.unpack_from('<H', entry, 64)[0]
        if entry[66] not in (2, 5) or not 2 <= name_len <= 64 or name_len % 2:
            continue
        name = entry[:name_len - 2].decode('utf-16le')
        value = (struct.unpack_from('<I', entry, 116)[0], struct.unpack_from('<Q', entry, 120)[0])
        if entry[66] == 5:
            root.append(value)
        elif name == 'Equation Native':
            found.append(value)
    if len(root) != 1 or len(found) != 1 or found[0][1] <= 28:
        raise ValueError('Equation Native 缺失或为空')
    first, length = found[0]
    if length < mini_cutoff:
        if minifat_count < 1:
            raise ValueError('迷你 FAT 缺失')
        raw_minifat = chain(first_minifat, fat, sector, minifat_count * size)
        minifat = struct.unpack('<' + 'I' * (len(raw_minifat) // 4), raw_minifat)
        mini_stream = chain(root[0][0], fat, sector, root[0][1])
        def mini_sector(n):
            offset = n * (1 << mini_shift)
            part = mini_stream[offset:offset + (1 << mini_shift)]
            if len(part) != (1 << mini_shift):
                raise ValueError('迷你流扇区越界')
            return part
        payload = chain(first, minifat, mini_sector, length)
    else:
        payload = chain(first, fat, sector, length)
    if not any(payload[28:]):
        raise ValueError('Equation Native 没有有效的非零内容')
    return len(payload)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('docx', type=Path)
    ap.add_argument('--expected-total', type=int, required=True, help='文档内预期的 MathType OLE 总数，包括原有对象')
    args = ap.parse_args()
    with zipfile.ZipFile(args.docx) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
        targets = {x.attrib['Id']: x.attrib['Target'] for x in rels}
        objects = [x for x in root.iter(OLE) if x.get('ProgID') == 'Equation.DSMT4']
        if len(objects) != args.expected_total:
            raise ValueError(f'MathType 对象数 {len(objects)}，预期 {args.expected_total}')
        sizes = []
        for obj in objects:
            target = targets[obj.attrib[REL]]
            path = posixpath.normpath(posixpath.join('word', target))
            if not path.startswith('word/embeddings/'):
                raise ValueError('公式不是文档内嵌对象')
            sizes.append(native_stream_size(z.read(path)))
    print(json.dumps({'objects': len(objects), 'equation_native_sizes': sizes,
                      'note': '仅检查对象结构和数据长度；必须另行检查公式内容、格式、编号和引用。'}, ensure_ascii=False))


if __name__ == '__main__':
    main()

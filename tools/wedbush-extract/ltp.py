"""LTP layer on top of pstlib: Heap-on-Node, BTree-on-Heap, Property Context."""
import struct
import datetime
from pstlib import PST, nid_type

FILETIME_EPOCH = datetime.datetime(1601, 1, 1, tzinfo=datetime.timezone.utc)


def filetime_to_dt(ft):
    if ft <= 0:
        return None
    try:
        return FILETIME_EPOCH + datetime.timedelta(microseconds=ft // 10)
    except OverflowError:
        return None


def hn_header_size(block_index):
    """Bytes of page header before user data in a given HN block."""
    if block_index == 0:
        return 12                      # HNHDR
    if block_index >= 8 and (block_index - 8) % 128 == 0:
        return 66                      # HNBITMAPHDR (ibHnpm + 64B fill level)
    return 2                           # HNPAGEHDR


class HN:
    """Heap-on-Node over a node's list of data blocks."""

    def __init__(self, blocks):
        self.blocks = blocks
        b0 = blocks[0]
        self.ibHnpm = struct.unpack_from('<H', b0, 0)[0]
        self.bSig = b0[2]
        self.bClientSig = b0[3]
        self.hidUserRoot = struct.unpack_from('<I', b0, 4)[0]
        if self.bSig != 0xEC:
            raise ValueError('bad HN signature 0x%02X' % self.bSig)

    def alloc(self, block_index, index):
        """Return bytes of allocation `index` (1-based) in block `block_index`."""
        blk = self.blocks[block_index]
        ibHnpm = struct.unpack_from('<H', blk, 0)[0]
        cAlloc, cFree = struct.unpack_from('<HH', blk, ibHnpm)
        if index < 1 or index > cAlloc:
            raise IndexError('alloc %d out of range (cAlloc=%d)' % (index, cAlloc))
        base = ibHnpm + 4
        start = struct.unpack_from('<H', blk, base + (index - 1) * 2)[0]
        end = struct.unpack_from('<H', blk, base + index * 2)[0]
        return blk[start:end]

    def get_hid(self, hid):
        if hid == 0:
            return None
        if (hid & 0x1F) != 0:
            raise ValueError('not a HID: 0x%08X' % hid)
        hidIndex = (hid >> 5) & 0x7FF
        hidBlockIndex = (hid >> 16) & 0xFFFF
        if hidIndex == 0:
            return None
        if hidBlockIndex >= len(self.blocks):
            raise IndexError('HN block %d >= %d' % (hidBlockIndex, len(self.blocks)))
        return self.alloc(hidBlockIndex, hidIndex)


class BTH:
    def __init__(self, hn, hidRoot):
        hdr = hn.get_hid(hidRoot)
        if hdr is None or len(hdr) < 8:
            raise ValueError('bad BTH header')
        self.bType, self.cbKey, self.cbEnt, self.bIdxLevels = hdr[0], hdr[1], hdr[2], hdr[3]
        self.hidRoot = struct.unpack_from('<I', hdr, 4)[0]
        if self.bType != 0xB5:
            raise ValueError('bad BTH bType 0x%02X' % self.bType)
        self.hn = hn

    def records(self):
        out = []
        if self.hidRoot:
            self._walk(self.hidRoot, self.bIdxLevels, out)
        return out

    def _walk(self, hid, level, out):
        data = self.hn.get_hid(hid)
        if not data:
            return
        if level > 0:
            rec = self.cbKey + 4
            for i in range(len(data) // rec):
                child = struct.unpack_from('<I', data, i * rec + self.cbKey)[0]
                self._walk(child, level - 1, out)
        else:
            rec = self.cbKey + self.cbEnt
            for i in range(len(data) // rec):
                off = i * rec
                out.append((data[off:off + self.cbKey], data[off + self.cbKey:off + rec]))


# fixed-size property types stored inline in dwValueHnid
INLINE_TYPES = {0x0002, 0x0003, 0x0004, 0x000A, 0x000B}
EIGHT_BYTE_TYPES = {0x0005, 0x0006, 0x0007, 0x0014, 0x0040}


class PropertyContext:
    def __init__(self, pst, nid, bids=None):
        """nid from the NBT, or bids=(bidData, bidSub) for a subnode."""
        self.pst = pst
        self.nid = nid
        if bids is None:
            bidData, bidSub, self.parent = pst.nbt[nid]
        else:
            bidData, bidSub = bids
            self.parent = None
        self.blocks = pst.data_blocks(bidData)
        self.hn = HN(self.blocks)
        if self.hn.bClientSig != 0xBC:
            raise ValueError('not a PC (bClientSig=0x%02X)' % self.hn.bClientSig)
        self.subs = pst.subnodes(bidSub)
        self.raw = {}
        for key, val in BTH(self.hn, self.hn.hidUserRoot).records():
            propid = struct.unpack('<H', key)[0]
            ptype, hnid = struct.unpack_from('<HI', val, 0)
            self.raw[propid] = (ptype, hnid)

    def _resolve(self, hnid):
        """HNID -> bytes (heap item, or concatenated subnode data)."""
        if hnid == 0:
            return None
        if (hnid & 0x1F) == 0:
            return self.hn.get_hid(hnid)
        ent = self.subs.get(hnid & 0xFFFFFFFF)
        if ent is None:
            return None
        return b''.join(self.pst.data_blocks(ent[0]))

    def get(self, propid, default=None):
        ent = self.raw.get(propid)
        if ent is None:
            return default
        ptype, hnid = ent
        if ptype in INLINE_TYPES:
            if ptype == 0x000B:
                return bool(hnid & 0xFFFF)
            if ptype == 0x0002:
                return struct.unpack('<h', struct.pack('<I', hnid)[:2])[0]
            if ptype == 0x0003:
                return struct.unpack('<i', struct.pack('<I', hnid))[0]
            if ptype == 0x0004:
                return struct.unpack('<f', struct.pack('<I', hnid))[0]
            return hnid
        data = self._resolve(hnid)
        if data is None:
            return default
        if ptype == 0x001F:
            return data.decode('utf-16-le', 'replace')
        if ptype == 0x001E:
            return data.decode('cp1252', 'replace')
        if ptype == 0x0040:
            return filetime_to_dt(struct.unpack('<Q', data[:8])[0]) if len(data) >= 8 else default
        if ptype == 0x0014:
            return struct.unpack('<q', data[:8])[0] if len(data) >= 8 else default
        if ptype == 0x0005:
            return struct.unpack('<d', data[:8])[0] if len(data) >= 8 else default
        return data


# property tags
P_SUBJECT = 0x0037
P_BODY = 0x1000
P_HTML = 0x1013
P_DELIVERY_TIME = 0x0E06
P_SUBMIT_TIME = 0x0039
P_SENDER_NAME = 0x0C1A
P_SENDER_EMAIL = 0x0C1F
P_NORMALIZED_SUBJECT = 0x0E1D
P_BODY_CODEPAGE = 0x3FDE
P_MSG_CLASS = 0x001A


def clean_subject(s):
    """Strip the PST subject-prefix control marker (0x01 + prefix length)."""
    if s and len(s) >= 2 and ord(s[0]) == 0x01:
        return s[2:]
    return s

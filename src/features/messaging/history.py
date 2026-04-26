# -*- coding: utf-8 -*-
"""消息历史存储。

为联系人/群聊提供按来源划分的、基于内存的消息历史保留。
向后的双端队列可防止在长时间运行的监听会话中出现无限制的内存增长。
"""

from __future__ import annotations

import time
from collections import deque
from typing import Deque, Dict, List, Optional


class MessageStore:
    """按联系人存储的消息历史。"""

    def __init__(self, max_per_contact: int = 500):
        self.max_per_contact = max_per_contact
        self._store: Dict[str, Deque[dict]] = {}

    def record(
        self,
        contact: str,
        content: str,
        *,
        sender: str = "",
        timestamp: Optional[float] = None,
        **extra,
    ) -> None:
        if not contact or not content:
            return

        entry = {
            "content": str(content),
            "sender": sender or "",
            "timestamp": timestamp or time.time(),
            **extra,
        }

        if contact not in self._store:
            self._store[contact] = deque(maxlen=self.max_per_contact)
        self._store[contact].append(entry)

    def get(self, contact: str, *, limit: int = 50) -> List[dict]:
        queue = self._store.get(contact)
        if not queue:
            return []
        items = list(queue)
        if limit > 0:
            items = items[-limit:]
        return items

    def last(self, contact: str) -> Optional[dict]:
        queue = self._store.get(contact)
        if not queue:
            return None
        return queue[-1]

    def all_contacts(self) -> List[str]:
        return sorted(self._store.keys())

    def clear(self, contact: Optional[str] = None) -> None:
        if contact:
            self._store.pop(contact, None)
        else:
            self._store.clear()

    def __contains__(self, contact: str) -> bool:
        return contact in self._store

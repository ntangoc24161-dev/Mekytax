# -*- coding: utf-8 -*-
"""
Tìm và phân tích các nhóm câu hỏi Anh em sinh đôi (Twin Questions)
"""
import openpyxl
import re
from collections import defaultdict

wb = openpyxl.load_workbook(r'D:\HOC TAP\ON THI 2026.09\40 DE THI\Tong_Hop_1429_Cau_Hoi_Doc_Ban_On_Tap_Nhanh.xlsx', data_only=True)
ws = wb['Ngan_Hang_1429_Cau_Hoi']
rows = list(ws.iter_rows(values_only=True))[1:]

questions = []
for r in rows:
    questions.append({
        'stt': r[0],
        'id': str(r[1]),
        'topic': r[2],
        'content': str(r[3]),
        'opt_a': str(r[4]),
        'opt_b': str(r[5]),
        'opt_c': str(r[6]),
        'opt_d': str(r[7]),
        'ans': str(r[8]),
        'ans_text': str(r[9]),
        'count': r[10],
        'exams': r[11],
        'note': str(r[12] or '')
    })

print(f"Tổng số câu hỏi độc bản: {len(questions)}")

def clean_tokens(text):
    text = re.sub(r'[^\w\s]', ' ', text.lower())
    return set([w for w in text.split() if len(w) > 1])

# Build inverted index
word_to_q = defaultdict(list)
for idx, q in enumerate(questions):
    tokens = clean_tokens(q['content'])
    for tok in tokens:
        word_to_q[tok].append(idx)

# Find candidate pairs
candidate_pairs = defaultdict(int)
for tok, q_indices in word_to_q.items():
    if len(q_indices) > 500: # skip overly common words
        continue
    for i in range(len(q_indices)):
        for j in range(i + 1, len(q_indices)):
            candidate_pairs[(q_indices[i], q_indices[j])] += 1

twins = []
for (i, j), common_cnt in candidate_pairs.items():
    q1 = questions[i]
    q2 = questions[j]
    t1 = clean_tokens(q1['content'])
    t2 = clean_tokens(q2['content'])
    union_len = len(t1.union(t2))
    jaccard = common_cnt / union_len if union_len > 0 else 0
    if jaccard >= 0.50:
        twins.append((jaccard, q1, q2))

twins.sort(key=lambda x: -x[0])
print(f"Tìm thấy {len(twins)} cặp câu hỏi có độ tương đồng Jaccard >= 0.50")

for jc, q1, q2 in twins[:25]:
    print("="*80)
    print(f"Jaccard: {jc:.2f} | Topic: {q1['topic']} vs {q2['topic']}")
    print(f"Q1 [ID {q1['id']} | Đáp án {q1['ans']}]: {q1['content']}")
    print(f"   -> Đúng: {q1['ans_text']}")
    print(f"Q2 [ID {q2['id']} | Đáp án {q2['ans']}]: {q2['content']}")
    print(f"   -> Đúng: {q2['ans_text']}")

import datetime
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

import openpyxl
import pyautogui
import pyperclip
from pydantic import BaseModel, Field

from auto_iuh import (
    cham_mot_cau_web,
    col_label,
    detect_columns_and_parse,
    detect_excel_schema,
    filter_start_students,
    format_score,
    normalize_mssv,
)

pyautogui.FAILSAFE = True

DEFAULT_MODEL = "gemini-3.8-flash"
DEFAULT_ARBITRATOR_MODEL = "gemini-3.7-flash"
AVAILABLE_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.1-pro-preview",
    "gemini-3-flash-preview",
    "gemini-4-argon",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-2.5-pro",
    "gemini-2.5-flash",
]
DEFAULT_THINKING_LEVEL = "high"
MAX_PROMPT_CLUSTERS_PER_PART = 2
MAX_ARBITRATOR_CASES_PER_CLASS = 6
BATCH_SIZE = 7
MAX_WORKERS = 2
_DAILY_EXHAUSTED_MODELS: set[str] = set()
_ACTIVE_THOUGHT_CALLBACK = None

JS_SCRAPER_CODE = r"""(async () => {
    console.log("🚀 Đang quét danh sách sinh viên trên bảng điểm Moodle...");
    const allRows = Array.from(document.querySelectorAll('table tbody tr'));
    const taskMap = new Map();

    for (const tr of allRows) {
        if (tr.classList.contains('emptyrow')) continue;
        const text = tr.innerText || "";
        const links = Array.from(tr.querySelectorAll('a[href*="reviewquestion.php"]'));

        // Tìm MSSV theo 3 tầng:
        // (1) sv_<MSSV> ở cột Mã số ID
        // (2) <MSSV>@iuh ở cột Địa chỉ thư điện tử (bắt được các bạn bị trống cột Mã số ID như Ngô Anh Thư)
        // (3) Ô Tên tài khoản chứa đúng 8 chữ số
        let mssv = null;
        const svMatch = text.match(/sv_(\d{6,12})/i);
        const emailMatch = text.match(/\b(\d{6,12})@iuh/i);
        if (svMatch) {
            mssv = svMatch[1];
        } else if (emailMatch) {
            mssv = emailMatch[1];
        } else {
            const cells = Array.from(tr.querySelectorAll('td, th')).map(td => (td.innerText || "").trim());
            const exactCell = cells.find(c => /^[12]\d{7}$/.test(c));
            if (exactCell) {
                mssv = exactCell;
            } else if (links.length > 0) {
                const fallbackMatch = text.match(/\b([12]\d{7})\b/);
                if (fallbackMatch) mssv = fallbackMatch[1];
            }
        }
        if (!mssv) continue;

        const q17 = links.find(a => a.href.includes('slot=17'))?.href || null;
        const q18 = links.find(a => a.href.includes('slot=18'))?.href || null;
        if (!q17 && !q18 && !svMatch && !emailMatch) continue;

        const userLinks = Array.from(tr.querySelectorAll('a[href*="/user/view.php"], a[href*="/user/profile.php"]'))
            .filter(a => !a.classList.contains('d-inline-block') && !a.querySelector('img, .userinitials'))
            .map(a => {
                const c = a.cloneNode(true);
                c.querySelectorAll('.userinitials, .accesshide, .sr-only, img').forEach(n => n.remove());
                return (c.textContent || "").trim();
            })
            .filter(t => t.length > 0 && !/^(xem|review|attempt|lịch sử)/i.test(t));
        const hoTen = Array.from(new Set(userLinks)).join(" ").trim();

        const existing = taskMap.get(mssv);
        if (!existing || (!existing.q17 && q17) || (!existing.q18 && q18)) {
            taskMap.set(mssv, { mssv, ho_ten: hoTen, q17, q18, p1: null, p2: null });
        }
    }

    const tasks = Array.from(taskMap.values());
    console.log(`✅ Tìm thấy ${tasks.length} sinh viên. Bắt đầu chạy 8 luồng hút Đề chuẩn & Bài làm (giữ nguyên 100% định dạng xuống dòng trên Moodle)...`);

    function htmlToRawLines(el, isPrompt = false) {
        if (!el) return [];
        const clone = el.cloneNode(true);
        const ownerDoc = clone.ownerDocument || document;
        clone.querySelectorAll('script, style, noscript').forEach(n => n.remove());

        const textarea = clone.querySelector('textarea');
        if (textarea && (textarea.value || textarea.textContent)) {
            return (textarea.value || textarea.textContent || "").replace(/\u00a0/g, ' ').split(/\r?\n/);
        }

        // Đối với vùng HTML (có các thẻ <p>, <div>, <br>, <span>... hoặc đề bài isPrompt = true):
        // Gộp các ký tự xuống dòng mềm (\n) nằm bên trong cùng 1 text node HTML thành khoảng trắng trước khi tách theo thẻ <p>, <div>, <br>
        const hasHtmlStructure = isPrompt || Boolean(clone.querySelector('p, div, br, span, ul, ol, li, table, b, i, u, strong, em, font'));
        if (hasHtmlStructure) {
            const walker = ownerDoc.createTreeWalker(clone, NodeFilter.SHOW_TEXT, null, false);
            let tNode;
            while ((tNode = walker.nextNode())) {
                if (tNode.parentElement && tNode.parentElement.closest('pre, textarea')) continue;
                tNode.nodeValue = (tNode.nodeValue || "").replace(/\s+/g, ' ');
            }
        }

        // Chuyển mọi thẻ <br> thành ký tự xuống dòng thật \n
        clone.querySelectorAll('br').forEach(br => {
            br.replaceWith(ownerDoc.createTextNode('\n'));
        });
        if (isPrompt) {
            clone.querySelectorAll('li').forEach(li => {
                li.insertBefore(ownerDoc.createTextNode('\n• '), li.firstChild);
                li.appendChild(ownerDoc.createTextNode('\n'));
            });
        }
        // Chèn \n vào đầu và cuối mọi thẻ khối (<p>, <div>, <li>, <tr>...) để không bị mất xuống dòng như khi dùng visibility:hidden
        clone.querySelectorAll('p, div, li, tr, h1, h2, h3, h4, h5, h6, blockquote, pre, hr').forEach(block => {
            block.insertBefore(ownerDoc.createTextNode('\n'), block.firstChild);
            block.appendChild(ownerDoc.createTextNode('\n'));
        });

        return (clone.textContent || "").replace(/\u00a0/g, ' ').split(/\r?\n/);
    }

    function extractPromptText(el) {
        if (!el) return "";
        const lines = htmlToRawLines(el, true)
            .map(line => line.replace(/[ \t]+/g, ' ').trim())
            .filter(line => line.length > 0 && !/^AVCB\.?\s*FINAL\s*WRITING\s*TEST$/i.test(line));
        const merged = [];
        for (const ln of lines) {
            if (merged.length > 0) {
                const prev = merged[merged.length - 1];
                const prevEndsOpen = !/[.?!,:;]$/.test(prev);
                const currIsHeaderOrBullet = /^(?:•|-|\*|PART\s*\d|Remember\b|Hi\b|Hello\b|Dear\b|Hey\b|Looking\s+forward|Best\b|Love\b)/i.test(ln);
                const prevEndsConnector = /\b(what|where|when|who|how|which|at\s+the|in\s+the|to\s+the|for\s+the|write|write\s+an|can\s+i|bring\s+some\s+food\s+or|take\s+us\s+to\s+the)$/i.test(prev);
                if (!currIsHeaderOrBullet && (prevEndsConnector || (prevEndsOpen && /^[a-z]/.test(ln)))) {
                    merged[merged.length - 1] = prev + " " + ln;
                    continue;
                }
            }
            merged.push(ln);
        }
        const deduped = [];
        for (const ln of merged) {
            if (deduped.length > 0) {
                const prevNorm = deduped[deduped.length - 1].replace(/^PART\s*\d+\s*:\s*/i, '').trim().toLowerCase();
                if (prevNorm === ln.toLowerCase()) continue;
            }
            deduped.push(ln);
        }
        return deduped.join('\n').trim();
    }

    function cleanMoodleTemplateDotsOnLine(line) {
        return (line || "")
            .replace(/\u00a0/g, ' ')
            // 1. Xóa mọi chuỗi từ 2 dấu chấm/gạch/ba chấm trở lên (dòng kẻ ............ mặc định của Moodle)
            .replace(/[._\-…]{2,}/g, ' ')
            // 2. Xóa dấu chấm/gạch đơn bị sót lại ở ngay đầu dòng (ví dụ ".Hi Nam" hoặc "..Hi Kiet" do xóa dòng chấm chưa hết)
            .replace(/^[._\-…]+\s*/, '')
            // 3. Xóa dấu chấm/gạch rời ở cuối dòng nếu đứng sau khoảng trắng
            .replace(/\s+[._\-…]+$/, '')
            .replace(/[ \t]+/g, ' ')
            .trim();
    }

    function extractEssayAndFormat(el) {
        if (!el) return { essay: "", essay_lines: [], word_count: 0, line_count: 0, format_type: "TRONG", format_note: "" };
        const rawLines = htmlToRawLines(el, false)
            .map(cleanMoodleTemplateDotsOnLine)
            .filter(line => line.length > 0);

        // Khử lỗi tự ý rớt dòng giữa câu do copy-paste từ đề bài hoặc HTML soft-wrap (VD: "what\ntime you go to bed", "how\noften you eat it", "at the\nrestaurant")
        const nonEmptyLines = [];
        for (let i = 0; i < rawLines.length; i++) {
            const ln = rawLines[i];
            if (nonEmptyLines.length > 0) {
                const prev = nonEmptyLines[nonEmptyLines.length - 1];
                const prevEndsOpen = !/[.?!,:;]$/.test(prev);
                const prevIsShortGreeting = /^(hi|hello|dear|hey|best|love|yours|your\s+friend|see\s+you|reply\s+(?:to\s+)?(?:me\s+)?soon|looking\s+forward)\b/i.test(prev) && (prev.match(/\S+/g) || []).length <= 5;
                const currIsSignoffOrGreeting = /^(hi|hello|dear|hey|best|love|yours|your\s+friend|see\s+you|reply\s+(?:to\s+)?(?:me\s+)?soon|looking\s+forward)\b/i.test(ln);
                const isLastSingleWordSignoff = (i === rawLines.length - 1) && (ln.match(/\S+/g) || []).length <= 2;
                const prevEndsMidPhrase = /\b(what|where|when|who|how|which|at\s+the|in\s+the|to\s+the|for\s+the|write|can\s+i|bring\s+some\s+food\s+or|wake\s+up\s+at|go\s+to)$/i.test(prev);
                if (!currIsSignoffOrGreeting && (prevEndsMidPhrase || (prevEndsOpen && !prevIsShortGreeting && !isLastSingleWordSignoff && /^[a-z]/.test(ln)))) {
                    nonEmptyLines[nonEmptyLines.length - 1] = prev + " " + ln;
                    continue;
                }
            }
            nonEmptyLines.push(ln);
        }

        const cleanEssay = nonEmptyLines.join('\n');
        const words = cleanEssay ? (cleanEssay.match(/[A-Za-z0-9]+(?:'[A-Za-z0-9]+)?/g) || []).length : 0;
        if (words === 0) {
            return { essay: "", essay_lines: [], word_count: 0, line_count: 0, format_type: "TRONG", format_note: "" };
        }

        // Phân tích phần thân bài xem sinh viên có bị lỗi "Xuống dòng rời rạc sau mỗi câu" không
        let bodyLines = nonEmptyLines.slice();
        const countLineWords = (s) => (s.match(/[A-Za-z0-9]+(?:'[A-Za-z0-9]+)?/g) || []).length;

        // Tách dòng chào đầu thư nếu ngắn (<= 6 từ)
        if (bodyLines.length > 1 && /^(hi|hello|dear|hey|good\s+(morning|afternoon|evening))\b/i.test(bodyLines[0]) && countLineWords(bodyLines[0]) <= 6) {
            bodyLines.shift();
        }
        // Tách 1-2 dòng kết thư / ký tên ở cuối nếu ngắn (<= 6 từ)
        if (bodyLines.length > 1 && countLineWords(bodyLines[bodyLines.length - 1]) <= 5) {
            bodyLines.pop();
        }
        if (bodyLines.length > 1 && /^(best|see\s+you|love|yours|regards|looking\s+forward|talk\s+to\s+you|write\s+soon|reply\s+soon|bye)\b/i.test(bodyLines[bodyLines.length - 1]) && countLineWords(bodyLines[bodyLines.length - 1]) <= 6) {
            bodyLines.pop();
        }

        const bodyWordCount = bodyLines.reduce((acc, l) => acc + countLineWords(l), 0);
        const avgWordsPerBodyLine = bodyLines.length > 0 ? (bodyWordCount / bodyLines.length) : 0;
        const maxBodyLineWords = bodyLines.length > 0 ? Math.max(...bodyLines.map(countLineWords)) : 0;

        let formatType = "BINH_THUONG";
        let formatNote = "";
        if (bodyLines.length >= 3 && avgWordsPerBodyLine <= 16 && maxBodyLineWords < 25) {
            formatType = "XUONG_DONG_TUNG_CAU";
            formatNote = `Sinh viên xuống dòng rời rạc sau mỗi câu đơn lẻ (${bodyLines.length} dòng thân bài) thay vì viết liền thành đoạn văn (trừ điểm ở tiêu chí Organization).`;
        }

        return {
            essay: cleanEssay,
            essay_lines: nonEmptyLines,
            word_count: words,
            line_count: nonEmptyLines.length,
            format_type: formatType,
            format_note: formatNote
        };
    }

    const imgCache = new Map();
    async function fetchImageAsBase64(imgUrl) {
        if (imgCache.has(imgUrl)) return imgCache.get(imgUrl);
        try {
            const resp = await fetch(imgUrl, { credentials: 'include' });
            if (!resp.ok) return null;
            const blob = await resp.blob();
            const dataUrl = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.readAsDataURL(blob);
            });
            const match = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
            const res = match ? { mime_type: match[1], data: match[2] } : null;
            imgCache.set(imgUrl, res);
            return res;
        } catch (err) {
            console.warn("⚠️ Không tải được ảnh đề thi:", imgUrl, err);
            return null;
        }
    }

    async function fetchQuestion(url) {
        if (!url) return { code: "DA_CHAM_HOAC_TRONG", prompt: "", images: [], essay: "", essay_lines: [], word_count: 0, line_count: 0, format_type: "TRONG", format_note: "" };
        const html = await (await fetch(url, { credentials: 'include' })).text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const title = (doc.querySelector('title')?.textContent || "").trim();
        const codeMatch = title.match(/AVCB[^\s]+/i);
        const qtextEl = doc.querySelector('.qtext');
        const prompt = extractPromptText(qtextEl);

        const images = [];
        if (qtextEl) {
            const imgEls = Array.from(qtextEl.querySelectorAll('img'));
            for (const img of imgEls) {
                const rawSrc = img.getAttribute('src');
                if (!rawSrc) continue;
                const absUrl = new URL(rawSrc, url).href;
                const imgObj = await fetchImageAsBase64(absUrl);
                if (imgObj) images.push(imgObj);
            }
        }

        const essayEl = doc.querySelector('.qtype_essay_response');
        const parsed = extractEssayAndFormat(essayEl);
        return {
            code: codeMatch ? codeMatch[0] : title,
            prompt,
            images,
            essay: parsed.essay,
            essay_lines: parsed.essay_lines,
            word_count: parsed.word_count,
            line_count: parsed.line_count,
            format_type: parsed.format_type,
            format_note: parsed.format_note
        };
    }

    let index = 0;
    async function worker(workerId) {
        while (index < tasks.length) {
            const current = index++;
            const item = tasks[current];
            item.p1 = await fetchQuestion(item.q17);
            item.p2 = await fetchQuestion(item.q18);
            console.log(`[Luồng ${workerId}] Đã lấy xong (${current + 1}/${tasks.length}): ${item.mssv} ${item.ho_ten} | P1: ${item.p1.word_count} từ, ${item.p1.line_count} dòng (${item.p1.format_type}) | P2: ${item.p2.word_count} từ, ${item.p2.line_count} dòng (${item.p2.format_type})`);
        }
    }

    await Promise.all(Array.from({ length: 8 }, (_, i) => worker(i + 1)));

    function detectClassCode() {
        const headerTexts = [
            document.title || "",
            ...Array.from(document.querySelectorAll('h1, h2, .breadcrumb, #page-navbar, .page-header-headings')).map(el => el.textContent || "")
        ].join(" ");
        const avcbMatch = headerTexts.match(/(\d{4}\s*AVCB\.?\s*\d{2,4})/i)
            || (document.body?.innerText || "").match(/(\d{4}\s*AVCB\.?\s*\d{2,4})/i);
        if (avcbMatch) {
            return avcbMatch[1].replace(/[\s.]+/g, '').toUpperCase();
        }
        const genericMatch = headerTexts.match(/(?:^|[^0-9A-Za-z])(\d{4}\s*[A-Z]{3,8}\.?\s*\d{2,4})(?=$|[^0-9A-Za-z])/i);
        if (genericMatch) {
            return genericMatch[1].replace(/[\s.]+/g, '').toUpperCase();
        }
        return "";
    }

    const classCode = detectClassCode();
    const fileName = classCode ? `bai_lam_iuh_${classCode}.json` : 'bai_lam_iuh.json';
    const blob = new Blob([JSON.stringify(tasks, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    a.click();
    console.log(`🎉 HOÀN TẤT! Đã tải file ${fileName} (${tasks.length} sinh viên) về máy!`);
})();"""

RUBRIC_VERSION_TAG = "QUY TẮC CHUẨN HÓA ĐIỂM GIẢNG VIÊN IUH - CUỐI KÌ SÁCH CAMBRIDGE (v7 - Calibrated on 2607AVCB005 & 2607AVCB006 Gold Standard)"

DEFAULT_RUBRIC = f"""# BỘ QUY TẮC CHẤM VIẾT: CUỐI KÌ — SÁCH CAMBRIDGE (CAMBRIDGE FINAL TERM)
**Môn học:** English for Beginners (AVCB — Sách Cambridge — Writing Part 1 & Part 2 Cuối Kì)  
**Phiên bản:** `{RUBRIC_VERSION_TAG}`  
**Thang điểm từng tiêu chí:** Số nguyên từ `0` đến `5` cho 4 cột: **Content**, **Organization**, **Grammar**, **Vocabulary**.  
*(Điểm từng Part = Trung bình cộng 4 cột `AVERAGE(C, O, G, V)`; Điểm Tổng toàn bài = `MROUND(Part 1 + Part 2, 0.5)`).*

> **NGUYÊN TẮC VIẾT NHẬN XÉT CHO GIẢNG VIÊN KIỂM CHỨNG (EXAMINER VERIFICATION FORMAT):**
> Phần nhận xét chi tiết (`rc, ro, rg, rv`) là để **Giảng viên chấm thi (Examiner)** đọc và kiểm chứng trong 2 giây.
> - **TUYỆT ĐỐI CẤM** viết nhận xét chung chung sáo rỗng (*"Đạt phần lớn yêu cầu, còn 1 chi tiết chưa rõ"*, *"Còn một số lỗi chia động từ, thiếu trợ động từ"*).
> - **BẮT BUỘC** trích dẫn trực tiếp cụm từ sai `"..."` trong bài sinh viên và chỉ đích danh ý nào của đề bị thiếu/lệch để Giảng viên đối chiếu ngay với bài làm.

---

## 1. TIÊU CHÍ CONTENT (0 – 5)

### 1.1. Nguyên tắc cốt lõi: Đặt mình vào vị trí Người Nhận Thư (Target Reader) & Kiểm tra đúng câu hỏi của Đề
* **Đếm ý linh hoạt theo đúng Đề bài thực tế của từng bài thi:**
  * Đọc kỹ từng gạch đầu dòng của `Remember to give information about:` và `Remember to ask about:` (ở Part 1) hoặc từng câu hỏi trong bức thư gốc (ở Part 2) để đối chiếu trực tiếp với bài làm.
* **Cách kiểm tra câu hỏi ở Part 1 (Tránh 2 bẫy phổ biến):**
  * **Bẫy 1 — Quên dấu chấm hỏi `?` hoặc dùng câu gián tiếp (KHÔNG trừ điểm Content):** Nếu sinh viên đã viết câu hỏi đúng chủ đề yêu cầu của đề nhưng kết thúc bằng dấu chấm `.` thay vì dấu `?` (ví dụ: *"What time do you wake up. And what your usually does in the morning."*, *"what your favourite food. Beaside, Where do you like eat it."*) hoặc dùng câu hỏi gián tiếp (*"Let me know what you do go after school, and what time you go to bed."*) $\\rightarrow$ **Tính là ĐÃ HỎI ĐỦ CÂU HỎI** (vẫn đạt `Content = 5` nếu đủ các ý khác; lỗi thiếu dấu `?` chỉ tính ở Organization/Grammar).
  * **Bẫy 2 — Dùng câu hỏi rủ rê/xã giao văn mẫu thay cho câu hỏi của đề (PHẢI trừ điểm Content):** Các câu hỏi xã giao/mời mọc như *"How are you?"*, *"Would you like to go to my house?"*, *"Is that OK for you?"*, *"OK for you?"*, *"Are you free?"*, *"Do you want to play with me?"* **KHÔNG** phải là câu hỏi chủ đề trong mục `Remember to ask about`. Nếu sinh viên chỉ có các câu hỏi mời mọc này mà không hỏi 2 câu hỏi của đề (ví dụ đề yêu cầu hỏi *where your friend lives* và *what his/her favorite room is*, nhưng chỉ viết *"Would you like to my house? ... OK for you?"*), hoặc biến 2 câu hỏi thành câu kể ngôi thứ 3 (*"My friend's favourite sport is badminton. She can play at school."*) $\\rightarrow$ Tính là **Quên toàn bộ phần câu hỏi cho bạn (`Content = 3`)**.
* **Cách kiểm tra vai giao tiếp ở Part 2 (3 mức độ Lộn vai / Đảo vai):**
  * Trong hầu hết các đề Part 2 (*"Thanks for inviting me to your party/picnic/cinema/museum... What time...? Can I bring...?"*), **bạn (người viết) là CHỦ TIỆC / NGƯỜI TỔ CHỨC**, còn người gửi thư là **KHÁCH ĐƯỢC MỜI**.
  * **Mức 1 — Lộn vai nhẹ ở câu chào đầu (`Content = 4`, trừ 1đ):** Câu đầu lỡ viết *"Would you like to come with me to the party?"* nhưng **cả 3 câu trả lời phía sau đều trả lời đúng vai chủ tiệc** (nêu rõ giờ bắt đầu, đồ mang theo, việc chuẩn bị).
  * **Mức 2 — Lộn vai câu đầu KÈM trả lời sai chủ thể hoặc ghép câu văn mẫu bận học thi (`Content = 3`, trừ 2đ):**
    * Mở đầu viết *"Thanks you for inviting."* VÀ khi khách hỏi *"Can I bring some food or drinks?"* lại trả lời ngược thành *"And I can bring some food and drinks."* (dùng `I` thay vì `you`);
    * Hoặc chép nguyên câu đầu của đề *"Thank you for inviting me to the cinema! I'd love to come."* rồi ghép câu văn mẫu *"I have a lesson at 9am to 10am for my exam on school. I free at 11am."* trước khi trả lời.
  * **Mức 3 — Đảo vai hoàn toàn / Ký tên bằng tên người gửi trên đề / Kể chuyện quá khứ / Trả lời cho Part 1 (`Content = 0`, khóa `0 - 1 - 1 - 1`):**
    * Chào tên mình và ký tên cuối thư bằng chính tên người gửi trên đề (ví dụ đề của `Sophie` lại chào `"Hi Dat,"` và ký tên `"Sophie"`) + ghép văn mẫu bận học thi;
    * Hoặc đề hỏi kế hoạch đi biển/cắm trại sắp tới nhưng lại viết đoạn văn kể chuyện quá khứ đã đi rồi (*"I went to Vung Tau with my friends..."*);
    * Hoặc ở Part 2 hoàn toàn phớt lờ đề bài Part 2 mà đóng vai người bạn ở Part 1 để viết thư trả lời cho bức thư của chính mình ở Part 1 (Tự biên tự diễn liên thông P1 $\\rightarrow$ P2).

### 1.2. Quy tắc Phần Kết Thư & Ký Tên (Closing & Sign-off)
* Cuối bức thư bắt buộc phải có **câu kết tạm biệt** (*Best, / See you soon, / Reply me soon,*) **VÀ ký tên riêng của người viết**.
* Nếu cuối bài dừng ngay sau câu hỏi cuối cùng, hoặc chỉ viết câu kết (*"See you soon!"*, *"Love you"*) mà không có tên, **HOẶC chỉ viết cụm danh xưng *"Your friend."* / *"Your best friend,"* / *"Best wishes,"* mà quên ghi tên riêng ở dưới**: **Trừ `1 điểm Content` (tối đa `Content = 4`)**.

### 1.3. Quy tắc Bài viết dài & Lan man (Không trừ điểm máy móc theo số từ)
* **KHÔNG trừ điểm máy móc chỉ vì bài dài > 100 từ:** Nếu bài viết dài 100–110 từ nhưng toàn bộ các câu đều tập trung trả lời chi tiết vào các yêu cầu của đề bài và hỏi đủ câu hỏi (không rào đón dông dài ngoài đề) $\\rightarrow$ **Vẫn đạt `Content = 5`**.
* **Chỉ trừ `1 điểm Content` (từ `5` $\\rightarrow$ `4`, hoặc từ `3` $\\rightarrow$ `2`)** khi bài viết dài ($\\ge 85$ từ) **VÀ thực sự LAN MAN**: có các câu rào đón sáo rỗng dài dòng (*"Long time no meet you, I miss you so much. Today I write an email to talk about my life / give information about... The first one... The last question..."*) hoặc kể lể chuyện ngoài lề không thuộc câu hỏi của đề. **Tuyệt đối không phạt dây chuyền sang `O, G, V`.**

### 1.4. Thang điểm Content chi tiết (0 – 5)
* **5 điểm:** Đáp ứng đầy đủ, rõ ràng 100% các ý thông tin & câu hỏi của đề (kể cả khi câu hỏi quên dấu `?` hoặc dùng câu gián tiếp *Let me know what...*), đúng vai người viết, có kết thư + ký tên riêng đầy đủ, không lan man.
* **4 điểm:**
  * Đủ 100% các ý của đề nhưng **thiếu kết thư hoặc thiếu ký tên riêng** (ví dụ kết bằng *"See you soon!"* hoặc *"Your friend."*);
  * Hoặc đủ các ý nhưng **viết dài lan man** (có câu rào đón/kể lể ngoài đề);
  * Hoặc đủ các ý nhưng giữ nguyên đại từ ngôi thứ 3 (*my friend / your friend / he / she*) trong câu hỏi cho bạn;
  * Hoặc **thiếu / diễn đạt hỏng 1 ý nhỏ** (ví dụ: ở Part 1 quên ý *how often you eat it*, hoặc chỉ hỏi 1/2 câu hỏi cho bạn; ở Part 2 lộn vai nhẹ ở câu mở đầu *"Would you like to come to the party?"* nhưng trả lời đúng cả 3 câu hỏi sau).
* **3 điểm:**
  * **Thiếu 2 ý / bỏ sót 2 câu hỏi** so với yêu cầu của đề (ví dụ: ở Part 1 kể đủ thông tin bản thân nhưng **quên đặt 2 câu hỏi cho bạn**, hoặc chỉ hỏi câu mời mọc *"Would you like to my house? OK for you?"*, hoặc viết 2 câu hỏi thành câu kể ngôi thứ 3; ở Part 2 bỏ sót câu hỏi phương tiện *"Is your mother going to take us to the park?"* + lan man/thiếu kết thư);
  * Hoặc ở Part 1 hỏi nhầm câu hỏi số 2 sang câu đã hỏi + giữ ngôi thứ 3 (*"What your friend's favourite food, Tom? How often you eat it?"* khi đề yêu cầu hỏi *where he/she likes to eat it*);
  * Hoặc ở Part 2 **mở đầu ngược vai (`"Thanks you for inviting"`) KÈM trả lời ngược chủ thể (`"I can bring some food"`) hoặc ghép câu văn mẫu bận học thi**.
* **2 điểm:**
  * **Trường hợp 2A (Không phạt dây chuyền):** Kể thông tin bản thân nhưng quên đặt 2 câu hỏi cho bạn **VÀ** thiếu cả chào/kết thư hoặc viết lan man (*ví dụ: viết thành 1 đoạn văn trần thuật không có mở/kết thư và không có câu hỏi*).
  * **Trường hợp 2B (Có phạt dây chuyền `-1` ở O, G, V):** Nửa đúng đề nửa học tủ/chép đề, hoặc chỉ viết câu cụt 18–25 từ trả lời 1 ý nhỏ và bỏ toàn bộ phần còn lại, hoặc biến toàn bộ gạch đầu dòng thông tin bản thân thành câu hỏi.
* **1 điểm (Không hiểu đề / Sai ngôi giao tiếp hoàn toàn — Phạt dây chuyền về `1 - 1 - 1 - 1` hoặc tối đa `1 - 2 - 2 - 2`):**
  * Tự đóng vai cả 2 người viết 2 bức thư đối đáp hỏi – trả lời trong cùng 1 bài làm;
  * Hoặc ở Part 1 biến cả bài thành thư rủ bạn đi chơi cuối tuần (không kể thói quen và không hỏi 2 câu của đề);
  * Hoặc ở Part 2 chép lại nguyên văn các câu hỏi của đề kèm vài từ vụn vặt, hoặc gọi người nhận thư ở ngôi thứ 3 (*"Tom will be come... Tom can bring..."*), hoặc từ chối lời mời vì bận thi rồi đổi sang ngày khác mà không trả lời các câu hỏi của đề.
* **0 điểm:**
  * **Lạc đề hoàn toàn / Học thuộc văn mẫu hẹn hò rủ rê (`Would you like to [eat/play/listen/go]... I'm free at... Let's meet at [23 Nguyen Van Bao / Circle K / Galaxy / Highlands]... Is that OK for you?`) / Đảo vai ký tên người gửi đề bài / Tự biên tự diễn P1-P2 liên thông:** **Khóa cứng `0 - 1 - 1 - 1`**.
  * **Bỏ trống, viết dở dang `< 12 từ` (hoặc `< 20 từ` chép văn mẫu quá khứ *It was a perfect trip*), hoặc chép nguyên văn đề bài:** **Khóa cứng `0 - 0 - 0 - 0` toàn bộ**.

---

## 2. TIÊU CHÍ ORGANIZATION (0 – 5)

### 2.1. Quy tắc kiểm soát lỗi Run-on Sentence / Comma Splice & Lặp cấu trúc `"I..."`
* **CẤM CHẤM 5 ĐIỂM ORGANIZATION** nếu bài viết mắc lỗi **Run-on sentence / Comma splice** (dùng dấu phẩy `,` thay dấu chấm `.` nối 2 mệnh đề độc lập không có liên từ, ví dụ: *"...have a pizza, rice, fish, My Mom is cooked it"*, *"I very happy, I am going to walk..."*).
  * Có 1–2 lỗi Comma splice: tối đa **`3 điểm`**.
  * Nối dấu phẩy tràn lan $\\ge 3$ chỗ: hạ xuống **`2 điểm`**.
* **Quy tắc Lặp cấu trúc đầu câu `"I / I am / I'm"`:**
  * Nếu lặp `"I / I am / I'm"` đầu câu từ 4 lần trở lên nhưng **có dùng từ nối** (*because, so, then, after that, besides...*): tối đa **`4 điểm`**.
  * Nếu lặp `"I..."` liên tiếp $\\ge 5$ câu đơn ngắn rời rạc mà **KHÔNG có từ nối chuyển ý nào** (như liệt kê từng câu *"I after go... I study... I go home... I sometime plays... I go to bed..."*): hạ `Organization` xuống **`3 điểm`** (không cho `4`).
* **Không trừ điểm khi trả lời linh hoạt thứ tự câu hỏi ở Part 2:**
  * Trong Part 2, nếu sinh viên trả lời câu mua vé trước rồi gộp giờ hẹn + quán cà phê vào câu sau để mạch văn tự nhiên hơn $\\rightarrow$ **Hoàn toàn hợp lý (`Organization = 5`)**, tuyệt đối không coi là lỗi ngược trình tự đề bài.

### 2.2. Quy tắc Trình bày (Xuống dòng rời rạc từng câu)
* Nếu sinh viên tự bấm `Enter` xuống dòng rời rạc sau mỗi câu đơn lẻ (biến thân bài thành danh sách 4–8 dòng ngắn): hạ `Organization` xuống **`3 điểm`** (hoặc **`2 điểm`** nếu các câu rời rạc không có từ nối).
* *(Lưu ý kỹ thuật: Không tính các chỗ ngắt dòng mềm do copy-paste giữa câu kiểu `"what\\ntime you"` là lỗi xuống dòng của sinh viên).*

---

## 3. TIÊU CHÍ GRAMMAR (0 – 5)

### 3.1. Nguyên tắc chấm Grammar & Trích dẫn lỗi kiểm chứng
* **Quy tắc chữ `i` viết thường / viết hoa nhẹ:** Bỏ qua lỗi viết thường chữ `i` đơn lẻ hoặc quên viết hoa chữ cái đầu câu 1 lần nếu cấu trúc ngữ pháp của câu hoàn toàn đúng (vẫn đạt `Grammar = 5` nếu không có lỗi cấu trúc nào khác).
* **5 điểm:** Cả bài kiểm soát tốt cấu trúc câu đơn/câu ghép và câu hỏi, không có lỗi cấu trúc ngữ pháp.
* **4 điểm:** Chỉ có **1 – 2 lỗi ngữ pháp nhẹ** (như thiếu mạo từ *a/the*, giới từ nhẹ, hoặc 1 lỗi chia động từ nhỏ).
* **3 điểm (Bắt buộc chấm `3` ngay từ đầu khi có $\\ge 2$ lỗi cấu trúc cơ bản):**
  * Ví dụ các lỗi cấu trúc cơ bản phải trích dẫn rõ:
    * Sai chia động từ ngôi thứ nhất/thứ ba: *"I sometime plays"*, *"I eats"*, *"I gets"*, *"I has"*, *"My house have"*, *"The party start"*, *"My mother cook"*;
    * Thừa/sai trợ động từ to be: *"I'm write"*, *"I am live"*, *"My Mom is cooked it"*, *"we are went"*, *"I don't can"*, *"I can to"*;
    * Thiếu động từ to be trước tính từ: *"It very delicous"*, *"I happy"*, *"I free at 11am"*, *"because big and beautiful"*;
    * Thiếu `to` / sai dạng động từ: *"like play"*, *"want eat"*, *"need bring"*, *"like eat"*;
    * Sai cấu trúc câu hỏi / thiếu trợ động từ `do/does`: *"What time you go"*, *"Where you go"*, *"what your usually does"*, *"What your does"*, *"What do you favourite food is"*, *"What you usually in the morning"*, *"Is that you ok?"*;
    * Sai cấu trúc cảm ơn/tính từ sở hữu: *"Thanks you"*, *"Thank you for your."*, *"My love my house"*, *"good for we"*, *"play with our"*.
* **1 – 2 điểm:** Sai cấu trúc ngữ pháp nặng ở hầu hết các câu trong bài (*"I'm like... I'm often play..."*, *"we're have... we're can... we're will..."*, *"Her very beautiful... buy clothers me"* $\\rightarrow$ `Grammar = 2`; câu hoàn toàn vỡ vụn *"I my are many is backrom"* $\\rightarrow$ `Grammar = 0` hoặc `1`).

---

## 4. TIÊU CHÍ VOCABULARY (0 – 5)

### 4.1. Phân biệt rõ "Từ vựng tự sản sinh tốt (4–5đ)" vs "Từ vựng cơ bản / Lặp đề (3đ)" vs "Từ vựng yếu & Sai chính tả nhiều (2đ)"
* **5 điểm:** Tự sản sinh được các cụm từ/collocation tự nhiên, cụ thể ngoài đề bài (ví dụ: *"wait in line"*, *"at the main entrance of the museum"*, *"just yourself"*, *"grab a quick bite"*, *"extended family"*, *"freshmen in university"*, *"walk around"*, *"fresh fruit and snacks"*, *"reduce stress and relax my mind"*) và không sai chính tả. **Cấm AI gán nhãn nhầm các cụm từ do sinh viên tự viết là "lặp lại đề bài"!**
* **4 điểm:** Dùng từ vựng đúng chủ đề, diễn đạt tự nhiên (*comfortable and chill*, *beef noodles soup, delicious*, *friendly, hardworking, kind and helpful*) hoặc có cụm từ tốt bù trừ cho 1 lỗi chính tả nhẹ.
* **3 điểm:** Chỉ dùng các từ vựng cơ bản quen thuộc (*bread, milk, rice, fish, house, room, happy, relax, park, food, drinks...*) hoặc lấy lại từ khóa từ đề bài; hoặc có 1–2 lỗi chính tả nhẹ (*delicous, beautifull, invation, restsurant*).
* **2 điểm:** Vốn từ nghèo nàn chỉ chép từ đề bài **KÈM nhiều lỗi chính tả/dùng từ sai nghĩa** (*"tow rooms, a bead room, My love my house, Reply soon me"*; *"intiving, beacause, oranger"*; *"vagetalbal, watcher"*).
"""


class KetQuaCham(BaseModel):
    nhan_xet: str = Field(
        default="",
        description=(
            "Phân tích lỗi tổng quát TRƯỚC KHI cho điểm (bằng tiếng Việt): "
            "P1 (số từ, đạt mấy trên tổng số ý yêu cầu của đề, lỗi chính); "
            "P2 (số từ, trả lời mấy trên tổng số câu hỏi của đề, lỗi chính)."
        ),
    )
    p1_content: int = Field(ge=0, le=5, description="Điểm Content Part 1 (0-5)")
    p1_org: int = Field(ge=0, le=5, description="Điểm Organization Part 1 (0-5)")
    p1_grammar: int = Field(ge=0, le=5, description="Điểm Grammar Part 1 (0-5)")
    p1_vocab: int = Field(ge=0, le=5, description="Điểm Vocabulary Part 1 (0-5)")
    p1_c_reason: str = Field(default="", description="Lý do điểm Content Part 1 (Đúng ý nào, thiếu/sai chỗ nào)")
    p1_o_reason: str = Field(default="", description="Lý do điểm Organization Part 1 (Liên kết câu, ngắt câu, xuống dòng)")
    p1_g_reason: str = Field(default="", description="Lý do điểm Grammar Part 1 (Cấu trúc đúng và các lỗi ngữ pháp cụ thể)")
    p1_v_reason: str = Field(default="", description="Lý do điểm Vocabulary Part 1 (Độ đa dạng từ vựng và lỗi chính tả/dùng từ)")
    p2_content: int = Field(ge=0, le=5, description="Điểm Content Part 2 (0-5)")
    p2_org: int = Field(ge=0, le=5, description="Điểm Organization Part 2 (0-5)")
    p2_grammar: int = Field(ge=0, le=5, description="Điểm Grammar Part 2 (0-5)")
    p2_vocab: int = Field(ge=0, le=5, description="Điểm Vocabulary Part 2 (0-5)")
    p2_c_reason: str = Field(default="", description="Lý do điểm Content Part 2 (Đúng ý nào, thiếu/sai chỗ nào)")
    p2_o_reason: str = Field(default="", description="Lý do điểm Organization Part 2 (Liên kết câu, ngắt câu, xuống dòng)")
    p2_g_reason: str = Field(default="", description="Lý do điểm Grammar Part 2 (Cấu trúc đúng và các lỗi ngữ pháp cụ thể)")
    p2_v_reason: str = Field(default="", description="Lý do điểm Vocabulary Part 2 (Độ đa dạng từ vựng và lỗi chính tả/dùng từ)")


class KetQuaChamSVTrongLo(BaseModel):
    mssv: str = Field(description="Mã số sinh viên giữ nguyên từ đề bài")
    nhan_xet: str = Field(
        default="",
        description="Tóm tắt ngắn gọn P1 và P2 bằng tiếng Việt.",
    )
    p1_content: int = Field(ge=0, le=5)
    p1_org: int = Field(ge=0, le=5)
    p1_grammar: int = Field(ge=0, le=5)
    p1_vocab: int = Field(ge=0, le=5)
    p1_c_reason: str = Field(default="", description="Lý do điểm Content P1 ngắn gọn (Đạt/Lỗi)")
    p1_o_reason: str = Field(default="", description="Lý do điểm Organization P1 ngắn gọn (Đạt/Lỗi)")
    p1_g_reason: str = Field(default="", description="Lý do điểm Grammar P1 ngắn gọn (Đạt/Lỗi)")
    p1_v_reason: str = Field(default="", description="Lý do điểm Vocabulary P1 ngắn gọn (Đạt/Lỗi)")
    p2_content: int = Field(ge=0, le=5)
    p2_org: int = Field(ge=0, le=5)
    p2_grammar: int = Field(ge=0, le=5)
    p2_vocab: int = Field(ge=0, le=5)
    p2_c_reason: str = Field(default="", description="Lý do điểm Content P2 ngắn gọn (Đạt/Lỗi)")
    p2_o_reason: str = Field(default="", description="Lý do điểm Organization P2 ngắn gọn (Đạt/Lỗi)")
    p2_g_reason: str = Field(default="", description="Lý do điểm Grammar P2 ngắn gọn (Đạt/Lỗi)")
    p2_v_reason: str = Field(default="", description="Lý do điểm Vocabulary P2 ngắn gọn (Đạt/Lỗi)")


class KetQuaChamLo(BaseModel):
    danh_sach: list[KetQuaChamSVTrongLo]


class KetQuaCham1PhanSV(BaseModel):
    mssv: str = Field(description="Mã số sinh viên giữ nguyên từ đề bài")
    content: int = Field(ge=0, le=5, description="Điểm Content của phần đang chấm (0-5)")
    org: int = Field(ge=0, le=5, description="Điểm Organization của phần đang chấm (0-5)")
    grammar: int = Field(ge=0, le=5, description="Điểm Grammar của phần đang chấm (0-5)")
    vocab: int = Field(ge=0, le=5, description="Điểm Vocabulary của phần đang chấm (0-5)")
    c_reason: str = Field(default="", description="Lý do điểm Content chi tiết kèm trích dẫn (Đạt/Lỗi)")
    o_reason: str = Field(default="", description="Lý do điểm Organization chi tiết kèm trích dẫn (Đạt/Lỗi)")
    g_reason: str = Field(default="", description="Lý do điểm Grammar chi tiết kèm trích dẫn cụm từ sai (Đạt/Lỗi)")
    v_reason: str = Field(default="", description="Lý do điểm Vocabulary chi tiết kèm trích dẫn (Đạt/Lỗi)")
    nhan_xet_ngan: str = Field(default="", description="Tóm tắt ngắn gọn 1 câu cho phần bài làm này")


class KetQuaChamCumMaDe(BaseModel):
    danh_sach: list[KetQuaCham1PhanSV]


class KetQuaTrongTaiSV(BaseModel):
    mssv: str = Field(description="Mã số sinh viên giữ nguyên từ đề bài")
    verdict: str = Field(
        default="Đã thẩm định 2 vòng",
        description="Quyết định phản biện của Trọng tài Vòng 2 (VD: 'Đồng thuận GK1' hoặc 'Điều chỉnh điểm' kèm 1 câu lập luận ngắn gọn)",
    )
    p1_content: int = Field(ge=0, le=5, description="Điểm chốt Content Part 1 (0-5)")
    p1_org: int = Field(ge=0, le=5, description="Điểm chốt Organization Part 1 (0-5)")
    p1_grammar: int = Field(ge=0, le=5, description="Điểm chốt Grammar Part 1 (0-5)")
    p1_vocab: int = Field(ge=0, le=5, description="Điểm chốt Vocabulary Part 1 (0-5)")
    p1_c_reason: str = Field(default="", description="Lý do Content P1 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p1_o_reason: str = Field(default="", description="Lý do Organization P1 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p1_g_reason: str = Field(default="", description="Lý do Grammar P1 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p1_v_reason: str = Field(default="", description="Lý do Vocabulary P1 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p2_content: int = Field(ge=0, le=5, description="Điểm chốt Content Part 2 (0-5)")
    p2_org: int = Field(ge=0, le=5, description="Điểm chốt Organization Part 2 (0-5)")
    p2_grammar: int = Field(ge=0, le=5, description="Điểm chốt Grammar Part 2 (0-5)")
    p2_vocab: int = Field(ge=0, le=5, description="Điểm chốt Vocabulary Part 2 (0-5)")
    p2_c_reason: str = Field(default="", description="Lý do Content P2 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p2_o_reason: str = Field(default="", description="Lý do Organization P2 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p2_g_reason: str = Field(default="", description="Lý do Grammar P2 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")
    p2_v_reason: str = Field(default="", description="Lý do Vocabulary P2 đã thẩm định kèm trích dẫn (Đạt/Lỗi)")


class KetQuaHoiDongTrongTai(BaseModel):
    nhan_xet_hoi_dong: str = Field(
        default="",
        description="Tóm tắt ngắn gọn kết luận phản biện của Trọng tài Vòng 2 đối với các bài nhạy cảm trong lớp.",
    )
    danh_sach: list[KetQuaTrongTaiSV]


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        parent_dir = os.path.dirname(exe_dir)
        if os.path.basename(exe_dir).lower() == "dist" and os.path.exists(os.path.join(parent_dir, "config.json")):
            return parent_dir
        return exe_dir
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_config_path() -> str:
    env_cfg = os.environ.get("IUH_SUITE_CONFIG_PATH", "").strip()
    if env_cfg:
        return env_cfg
    return os.path.join(get_app_dir(), "config.json")


DEFAULT_RULE_PROFILE = "CAMBRIDGE_CUOI_KI"


def _build_exam_profile_template(title: str, subtitle: str, version_tag: str) -> str:
    return f"""# {title}
**Môn học:** {subtitle}  
**Phiên bản:** `{version_tag}`  
**Thang điểm từng tiêu chí:** Số nguyên từ `0` đến `5` cho 4 cột: **Content**, **Organization**, **Grammar**, **Vocabulary**.  
*(Điểm từng Part = Trung bình cộng 4 cột `AVERAGE(C, O, G, V)`; Điểm Tổng toàn bài = `MROUND(Part 1 + Part 2, 0.5)`).*

> **NGUYÊN TẮC VIẾT NHẬN XÉT CHO GIẢNG VIÊN KIỂM CHỨNG (EXAMINER VERIFICATION FORMAT):**
> Phần nhận xét chi tiết (`rc, ro, rg, rv`) là để **Giảng viên chấm thi (Examiner)** đọc và kiểm chứng trong 2 giây.
> - **TUYỆT ĐỐI CẤM** viết nhận xét chung chung sáo rỗng.
> - **BẮT BUỘC** trích dẫn trực tiếp cụm từ sai `"..."` trong bài sinh viên và chỉ đích danh ý nào của đề bị thiếu/lệch để Giảng viên đối chiếu ngay với bài làm.

---

## 1. TIÊU CHÍ CONTENT (0 – 5)
* **5 điểm:** Đáp ứng đầy đủ 100% các yêu cầu thông tin & câu hỏi của đề bài, đúng vai giao tiếp, có mở/kết đầy đủ, không lan man.
* **4 điểm:** Thiếu 1 chi tiết nhỏ hoặc thiếu kết thư/ký tên hoặc viết dài lan man ngoài đề (-1đ).
* **3 điểm:** Thiếu 2 ý / bỏ sót câu hỏi yêu cầu của đề hoặc lộn vai giao tiếp (-2đ).
* **2 điểm:** Chỉ trả lời được 1 phần nhỏ yêu cầu của đề hoặc chép lại nhiều từ của đề bài.
* **1 điểm:** Không hiểu đề / sai ngôi giao tiếp hoàn toàn (phạt dây chuyền tối đa `1 - 2 - 2 - 2` hoặc `1 - 1 - 1 - 1`).
* **0 điểm:** Lạc đề hoàn toàn / học thuộc văn mẫu không khớp đề (**khóa cứng `0 - 1 - 1 - 1`**); bỏ trống hoặc `< 12 từ` hoặc chép nguyên văn đề bài (**khóa cứng `0 - 0 - 0 - 0`**).

## 2. TIÊU CHÍ ORGANIZATION (0 – 5)
* **5 điểm:** Bố cục rõ ràng, mạch lạc, dùng từ nối tự nhiên, ngắt câu chuẩn xác.
* **4 điểm:** Bố cục hợp lý nhưng lặp cấu trúc đầu câu `"I..."` từ 4 lần trở lên (có từ nối).
* **3 điểm:** Mắc lỗi Comma splice (dùng dấu phẩy nối câu độc lập), hoặc xuống dòng rời rạc từng câu đơn lẻ, hoặc lặp $\\ge 5$ câu `"I..."` không có từ nối.
* **1 – 2 điểm:** Câu văn rời rạc, nối dấu phẩy tràn lan $\\ge 3$ chỗ hoặc thiếu mạch lạc nghiêm trọng.

## 3. TIÊU CHÍ GRAMMAR (0 – 5)
* **5 điểm:** Kiểm soát tốt cấu trúc câu, không có lỗi ngữ pháp cấu trúc (bỏ qua lỗi viết thường chữ `i` đơn lẻ).
* **4 điểm:** Chỉ có 1 lỗi ngữ pháp nhẹ (mạo từ, giới từ, chia động từ nhẹ).
* **3 điểm:** Có $\\ge 2$ lỗi cấu trúc cơ bản (sai chia động từ, thừa/thiếu to be, thiếu trợ động từ trong câu hỏi, thiếu `to` giữa 2 động từ) — bắt buộc trích dẫn rõ từng cụm sai.
* **1 – 2 điểm:** Lỗi cấu trúc câu nặng ở hầu hết các câu trong bài.

## 4. TIÊU CHÍ VOCABULARY (0 – 5)
* **5 điểm:** Tự sản sinh được các cụm từ/collocation tự nhiên ngoài đề bài, không sai chính tả.
* **4 điểm:** Từ vựng đúng chủ đề, tự nhiên hoặc có cụm từ tốt bù trừ 1 lỗi chính tả nhẹ.
* **3 điểm:** Từ vựng ở mức cơ bản quen thuộc hoặc lặp lại từ khóa từ đề bài, hoặc có 1–2 lỗi chính tả nhẹ.
* **1 – 2 điểm:** Vốn từ quá hạn chế kèm nhiều lỗi chính tả/dùng từ sai nghĩa.
"""


MARKING_RULE_PROFILES: dict[str, dict[str, str]] = {
    "CAMBRIDGE_CUOI_KI": {
        "label": "Cuối kì - Sách Cambridge",
        "filename": "MARKING_RULES_CAMBRIDGE_CUOI_KI.md",
        "version_tag": RUBRIC_VERSION_TAG,
        "default_content": DEFAULT_RUBRIC,
    },
    "CAMBRIDGE_GIUA_KI": {
        "label": "Giữa kì - Sách Cambridge",
        "filename": "MARKING_RULES_CAMBRIDGE_GIUA_KI.md",
        "version_tag": "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - GIỮA KÌ SÁCH CAMBRIDGE (v1)",
        "default_content": _build_exam_profile_template(
            "BỘ QUY TẮC CHẤM VIẾT: GIỮA KÌ — SÁCH CAMBRIDGE (CAMBRIDGE MIDTERM)",
            "English for Beginners (Sách Cambridge — Writing Giữa Kì)",
            "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - GIỮA KÌ SÁCH CAMBRIDGE (v1)",
        ),
    },
    "OXFORD_GIUA_KI": {
        "label": "Giữa kì - Sách Oxford",
        "filename": "MARKING_RULES_OXFORD_GIUA_KI.md",
        "version_tag": "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - GIỮA KÌ SÁCH OXFORD (v1)",
        "default_content": _build_exam_profile_template(
            "BỘ QUY TẮC CHẤM VIẾT: GIỮA KÌ — SÁCH OXFORD (OXFORD MIDTERM)",
            "English (Sách Oxford — Writing Giữa Kì)",
            "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - GIỮA KÌ SÁCH OXFORD (v1)",
        ),
    },
    "OXFORD_CUOI_KI": {
        "label": "Cuối kì - Sách Oxford",
        "filename": "MARKING_RULES_OXFORD_CUOI_KI.md",
        "version_tag": "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - CUỐI KÌ SÁCH OXFORD (v1)",
        "default_content": _build_exam_profile_template(
            "BỘ QUY TẮC CHẤM VIẾT: CUỐI KÌ — SÁCH OXFORD (OXFORD FINAL TERM)",
            "English (Sách Oxford — Writing Cuối Kì)",
            "QUY TẮC CHẤM VIẾT GIẢNG VIÊN IUH - CUỐI KÌ SÁCH OXFORD (v1)",
        ),
    },
}


def resolve_rule_profile_key(profile_key: str | None = None) -> str:
    if profile_key and profile_key in MARKING_RULE_PROFILES:
        return profile_key
    cfg_prof = str(load_config().get("rule_profile") or "").strip()
    if cfg_prof in MARKING_RULE_PROFILES:
        return cfg_prof
    return DEFAULT_RULE_PROFILE


def get_marking_rules_path(profile_key: str | None = None) -> str:
    key = resolve_rule_profile_key(profile_key)
    filename = MARKING_RULE_PROFILES[key]["filename"]
    return os.path.join(get_app_dir(), filename)


def load_marking_rules(profile_key: str | None = None) -> str:
    """Đọc bộ quy tắc chấm theo đợt thi & bộ sách (mặc định: MARKING_RULES_CAMBRIDGE_CUOI_KI.md)."""
    key = resolve_rule_profile_key(profile_key)
    prof = MARKING_RULE_PROFILES[key]
    rules_path = get_marking_rules_path(key)
    version_tag = prof["version_tag"]
    default_content = prof["default_content"]

    # Tự động dọn file tên cũ MARKING_RULES_IUH.md nếu còn tồn tại để tránh nhầm lẫn
    legacy_path = os.path.join(get_app_dir(), "MARKING_RULES_IUH.md")
    if os.path.exists(legacy_path) and os.path.abspath(legacy_path) != os.path.abspath(rules_path):
        try:
            os.remove(legacy_path)
        except Exception:
            pass

    if os.path.exists(rules_path):
        try:
            with open(rules_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            if content and version_tag in content:
                return content
        except Exception:
            pass
    try:
        with open(rules_path, "w", encoding="utf-8") as f:
            f.write(default_content)
    except Exception:
        pass
    return default_content


def load_config() -> dict:
    cfg_path = get_config_path()
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(data: dict):
    cfg_path = get_config_path()
    cfg = load_config()
    cfg.update(data)
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def clean_prompt_text(prompt: str) -> str:
    """Làm sạch mã JS thừa, nối các dòng bị ngắt gãy giữa câu trong HTML Moodle và khử dòng tiêu đề lặp."""
    if not prompt:
        return ""
    s = re.sub(r"document\.querySelectorAll\([\s\S]*$", "", prompt).replace("\r\n", "\n").strip()
    if "\n\n" in s:
        blocks = [b.strip() for b in re.split(r"\n\s*\n+", s) if b.strip()]
        joined_blocks: list[str] = []
        for b in blocks:
            sublines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in b.split("\n") if ln.strip()]
            merged: list[str] = []
            for ln in sublines:
                if (
                    not merged
                    or ln.startswith(("•", "-", "*"))
                    or re.match(r"^(PART\s+\d|Remember\b|Hi\b|Hello\b|Hey\b|Dear\b)", ln, re.IGNORECASE)
                ):
                    merged.append(ln)
                else:
                    merged[-1] = merged[-1] + " " + ln
            joined_blocks.extend(merged)
        lines = joined_blocks
    else:
        raw_lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n") if ln.strip()]
        lines = []
        for ln in raw_lines:
            if (
                lines
                and not lines[-1].endswith((".", "?", "!", ":", ","))
                and not ln.startswith(("•", "-", "*"))
                and re.match(r"^[a-z]", ln)
            ):
                lines[-1] = lines[-1] + " " + ln
            else:
                lines.append(ln)

    out: list[str] = []
    for ln in lines:
        if re.match(r"^AVCB\.?\s*FINAL\s*WRITING\s*TEST$", ln, re.IGNORECASE):
            continue
        if out and re.sub(r"^PART\s*\d+\s*:\s*", "", out[-1], flags=re.IGNORECASE).strip().lower() == ln.lower():
            continue
        out.append(ln)
    return "\n".join(out).strip()


def _extract_prompt_info(prompt_text: str, is_part1: bool) -> dict:
    """Trích xuất danh sách các ý/câu hỏi yêu cầu và tên người gửi từ Đề bài để đối chiếu chi tiết."""
    clean_p = clean_prompt_text(prompt_text)
    lines = [ln.strip() for ln in clean_p.splitlines() if ln.strip()]
    info_bullets: list[str] = []
    ask_bullets: list[str] = []
    p2_questions: list[str] = []
    sender_name = ""

    if is_part1:
        section = "info"
        for ln in lines:
            low = ln.lower()
            if "remember to ask" in low:
                section = "ask"
                continue
            if "remember to give" in low:
                section = "info"
                continue
            if ln.startswith(("•", "-", "*")):
                item = ln.lstrip("•-* ").strip()
                if section == "ask":
                    ask_bullets.append(item)
                else:
                    info_bullets.append(item)
    else:
        for ln in lines:
            if re.match(r"^PART\s*2\b", ln, re.IGNORECASE):
                continue
            for q_match in re.findall(r"[^.?!]*\?", ln):
                q_clean = q_match.strip()
                if q_clean:
                    p2_questions.append(q_clean)
            if re.search(r"\blet\s+me\s+know\s+[^.?!]+[.!]", ln, re.IGNORECASE):
                for imp_match in re.findall(r"[^.?!]*\blet\s+me\s+know\s+[^.?!]+[.!]", ln, re.IGNORECASE):
                    p2_questions.append(imp_match.strip())
        if lines:
            last_ln = lines[-1].strip(" .,!")
            if re.match(r"^[A-Z][a-z]{1,15}$", last_ln) and last_ln.lower() not in ("hi", "hey", "hello", "best", "love"):
                sender_name = last_ln

    return {
        "info_bullets": info_bullets,
        "ask_bullets": ask_bullets,
        "p2_questions": p2_questions,
        "sender_name": sender_name,
    }


def count_words(text: str) -> int:
    s = (text or "").strip()
    if not s:
        return 0
    # Tách thêm các chỗ bị dính dấu chấm/phẩy như "Hi Mary,To" -> đếm đúng từ thực tế
    tokens = re.findall(r"[A-Za-z0-9]+(?:'[A-Za-z0-9]+)?", s)
    return len(tokens)


def _count_non_copied_words(prompt_text: str, essay_text: str) -> tuple[int, float]:
    """Đếm số từ phần thân bài do sinh viên tự viết (loại bỏ chào/kết thư và các cụm >= 4 từ chép nguyên văn từ đề bài)."""
    lines = [ln.strip() for ln in (essay_text or "").splitlines() if ln.strip()]
    body_lines = []
    for idx, ln in enumerate(lines):
        if idx == 0 and re.match(r"^(hi|hello|dear|hey)\b", ln, re.IGNORECASE) and count_words(ln) <= 4:
            continue
        if idx >= len(lines) - 2 and re.match(
            r"^(see\s+you|best|love|bye|yours|looking\s+forward|reply\s+soon)", ln, re.IGNORECASE
        ) and count_words(ln) <= 5:
            continue
        if idx == len(lines) - 1 and count_words(ln) <= 2:
            continue
        body_lines.append(ln)
    body_text = " ".join(body_lines) if body_lines else (essay_text or "")
    essay_words = re.findall(r"[a-z0-9']+", body_text.lower())
    if not essay_words:
        return 0, 0.0
    prompt_clean = clean_prompt_text(prompt_text).lower()
    prompt_words = re.findall(r"[a-z0-9']+", prompt_clean)
    if len(prompt_words) < 4 or len(essay_words) < 4:
        return len(essay_words), 0.0

    prompt_ngrams = {tuple(prompt_words[i:i + 4]) for i in range(len(prompt_words) - 3)}
    copied_flags = [False] * len(essay_words)
    for i in range(len(essay_words) - 3):
        if tuple(essay_words[i:i + 4]) in prompt_ngrams:
            for k in range(i, i + 4):
                copied_flags[k] = True

    copied_count = sum(1 for c in copied_flags if c)
    own_count = len(essay_words) - copied_count
    ratio = copied_count / len(essay_words)
    return own_count, ratio


_COMMON_MISSPELLINGS = re.compile(
    r"\b(talking\s+photos?|wil|verry|fovourite|fovorite|faverite|paceful|cenima|ciname|strees|busetyful|"
    r"beautifull|beautyfull|futer|coffer|coffe|vestables|laughting|becaues|beacause|wuold|manny|"
    r"frienly|saturdsy|dirnks|shoses|persent|bancol|whies|lfavourite|weakeend|invation|"
    r"popoc|feld|okey|gong|a\s+bout|let\s+me\s+now|delicous|deliciuos|beaside|remmember|comfortble|"
    r"usuall|tow\s+rooms|bead\s+room|badmintion|badminon|footbal|umbrela|umbrellar|sandwitch|"
    r"musium|muesum|exited|tomorow|restarant|restautant|studing|swimming\s+poll|aprtment|aparment|pictic|"
    r"monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b",
    re.IGNORECASE,
)

_GOOD_VOCAB_PHRASES = re.compile(
    r"\b(big\s+and\s+busy|shops\s+and\s+restaurants|curious|stay\s+healthy|"
    r"walk\s+around|fresh\s+fruits?|fried\s+chicken|pink\s+dress|black\s+coat|"
    r"for\s+this\s+winter|delicious|reasonable\s+price|reduce\s+stress|"
    r"relax\s+my\s+mind|comfortable\s+clothes|sneakers|am\s+very\s+excited|"
    r"wait\s+in\s+line|main\s+entrance|fresh\s+air|spend\s+time\s+with|comfortable\s+and\s+chill)\b",
    re.IGNORECASE,
)

def _count_comma_splices_and_fragments(essay_text: str) -> int:
    """Đếm số lỗi dùng dấu phẩy ',' thay dấu chấm câu '.' nối 2 mệnh đề độc lập (Run-on sentence / Comma splice)."""
    if not essay_text:
        return 0
    count = 0
    intro_prefix = re.compile(
        r"^(?:hi|hello|dear|hey|good\s+\w+|to\s+\w+|for\s+\w+|if\b|when\b|while\b|because\b|as\b|since\b|"
        r"although\b|after\b|before\b|in\b|on\b|at\b|sometimes\b|usually\b|often\b|first\b|second\b|"
        r"moreover\b|however\b|about\b|and\s+you\b|oh\b|yes\b|no\b|ok\b|sorry\b|thank|see\s+you|best|love|reply)",
        re.IGNORECASE,
    )
    next_clause_start = re.compile(
        r"^(?:i\b|i'm\b|we\b|we're\b|you\b|he\b|she\b|it\b|it's\b|they\b|there\s+(?:is|are)\b|"
        r"is\s+that\b|can\s+you\b|let's\b|my\s+\w+\s+(?:is|are|was|were|cooked)\b|hope\s+will\b|6:30\b)",
        re.IGNORECASE,
    )
    for raw_line in essay_text.splitlines():
        for sent in re.split(r"[.?!]+", raw_line):
            parts = [p.strip() for p in sent.split(",") if p.strip()]
            if len(parts) < 2:
                continue
            for idx in range(len(parts) - 1):
                next_p = parts[idx + 1]
                has_prior_clause = any(
                    count_words(p) >= 3 and not intro_prefix.match(p)
                    for p in parts[: idx + 1]
                )
                if has_prior_clause and next_clause_start.match(next_p):
                    count += 1
    return count

_SEVERE_GRAMMAR_ERRORS = re.compile(
    r"\b(i\s+am\s+(?:live|go|play|eat|study|wake|get|have|sleep|watch|listen|walk|cook|buy|bring|meet|take|visit|stay|work|do|make|read|write|drink)|"
    r"i'm\s+(?:live|go|play|eat|study|wake|get|have|sleep|watch|listen|walk|cook|buy|bring|meet|take|visit|stay|work|do|make|read|write|drink|like|love|often|usually|always|sometimes)|"
    r"we\s+can\s+going|we're\s+(?:have|can|will)|we\s+are\s+(?:plays|went|go)|"
    r"i\s+don't\s+can\b|you\s+are\s+go\b|are\s+you\s+go\b|do\s+you\s+free\b|"
    r"i'm\s+interesting\s+my\s+party|i\s+only\s+emma\s+is|i'd\s+love\s+i'm\s+free)\b",
    re.IGNORECASE,
)

_COMMON_GRAMMAR_ERRORS = re.compile(
    r"\b(i'am|(?:like|love|enjoy)\s+(?:play|eat|go|watch|listen|read|cook|swim|walk|drink|do|visit|stay|study|buy|make)\b|"
    r"(?:want|need)\s+(?:eat|go|bring|play|watch|come|meet|buy|visit|drink|see|take|have)\b|"
    r"my\s+mom\s+is\s+cooked|is\s+cooked\s+it|i\s+have\s+breakfast\s+is|my\s+love\s+my\b|i\s+after\s+go\b|what\s+your\s+usually\s+does|"
    r"(?:i|we|you|they)\s+(?:plays|eats|goes|likes|loves|wants|needs|has|watches|studies|drinks|wakes|gets|does|cooks|takes|makes|lives|works)\b|"
    r"i\s+sometime\s+plays|sometime\s+plays|"
    r"(?:he|she|it|my\s+(?:house|mother|father|mom|dad|friend|sister|brother|teacher|party|school|room|city|hometown))\s+(?:have|go|play|eat|like|love|start|end|begin|finish|cook|make|take)\b|"
    r"(?:it|he|she|my\s+\w+)\s+very\s+(?:delicous|delicious|good|nice|big|small|happy|beautiful|peaceful|fun)\b|"
    r"i\s+going\s+to|we\s+going\s+to|you\s+going\s+to|is\s+have\s+fun|"
    r"places\s+is|my\s+many\s+shops|their\s+are|what\s+your\s+(?:favourite|favorite|name|hobby|food|sport|room)\b|"
    r"where\s+you\s+(?:go|live|eat|play|like|want|study|stay)|what\s+you\s+(?:do|eat|play|like|want|have|study)|"
    r"what\s+time\s+you\s+(?:go|wake|get|eat|have|sleep|start|meet|come|can)|"
    r"how\s+often\s+you\s+(?:go|eat|play|do|have|visit|watch)|who\s+you\s+(?:go|play|eat|live)|"
    r"what\s+do\s+you\s+go\b|what\s+you\s+do\s+go\b|do\s+you\s+want\s+to\s+anything|would\s+you\s+like\s+to\s+my\s+house|"
    r"i\s+am\s+free\s+time|is\s+that\s+you\s+ok|it\s+that\s+ok|it's\s+that\s+ok|"
    r"thank\s+to\b|thanks\s+you\b|thank\s+your\b|thank\s+you\s+for\s+reply\b|"
    r"look\s+forward\s+to\s+(?:answering\s+by\s+you|see\s+you|meet\s+you|hear\s+from\s+you)|"
    r"reply\s+soon\s+me\b|go\s+to\s+the\s+bed\b|in\s+weekend\b|in\s+saturday\b|in\s+sunday\b|"
    r"can\s+you\s+watch\s+the\s+movie\s+title|do\s+not\s+late|hope\s+emma\s+not\s+come|"
    r"a\s+coffee\s+so\s+good|a\s+drink\s+is\b|have\s+a\s+drink\s+\w+\s+juice|the\s+party\s+is\s+start|"
    r"what\s+you\s+likes|a\s+new\s+clothes|a\s+bread\b|a\s+milk\b|a\s+rice\b|a\s+food\b|"
    r"i\s+very\s+(?:happy|excited|miss|like|love|glad))\b",
    re.IGNORECASE,
)


def _missing_signoff_or_closing(essay_text: str) -> bool:
    """Kiểm tra xem cuối bức thư có bị thiếu kết thư hoặc thiếu ký tên người viết không (bao gồm cả lỗi chỉ ghi 'Your friend.' / 'Dear,' mà không có tên)."""
    s = (essay_text or "").strip()
    if not s:
        return True
    if s.endswith("?"):
        return True
    tail = s[-55:].strip()
    if re.search(
        r"(?:see\s+you(?:\s+soon|\s+later|\s+on\s+\w+|\s+at\s+\w+|\s+this\s+\w+)?|"
        r"looking\s+forward\s+to\s+[\w\s]+|reply(?:\s+to)?\s+me\s+soon|write\s+soon|"
        r"hope\s+to\s+see\s+you(?:\s+soon)?|can't\s+wait|love\s+you|"
        r"your\s+(?:best\s+)?friend|best\s+wishes|best\s+regards|warm\s+regards|yours|best|love|dear)[!.,]*$",
        tail,
        re.IGNORECASE,
    ):
        return True
    return False


def _extract_greeting_and_signoff_names(essay_text: str) -> tuple[str, str]:
    """Trích xuất tên người nhận ở dòng chào đầu thư ('Hi Hung,') và tên người ký ở dòng cuối thư ('Truong')."""
    lines = [ln.strip() for ln in (essay_text or "").splitlines() if ln.strip()]
    if not lines:
        return "", ""
    greet_name = ""
    m_greet = re.match(r"^(?:hi|hello|dear|hey)\s*,?\s*([A-Za-z]+)\b", lines[0], re.IGNORECASE)
    if m_greet:
        cand = m_greet.group(1).strip().lower()
        if cand not in ("my", "friend", "friends", "there", "all", "everyone", "teacher", "how"):
            greet_name = cand
    sign_name = ""
    last_ln = re.sub(r"[.!?,]+$", "", lines[-1]).strip()
    if re.match(r"^[A-Za-z]{2,15}$", last_ln):
        cand_s = last_ln.lower()
        if cand_s not in ("best", "love", "yours", "dear", "bye", "thanks", "thank", "friend", "soon", "me", "you"):
            sign_name = cand_s
    return greet_name, sign_name


def _has_rambling_or_meta_filler(prompt_text: str, essay_text: str, is_part1: bool) -> bool:
    """Kiểm tra xem bài viết dài có bị lan man ngoài đề, chêm câu dẫn nhập sáo rỗng hoặc kể lể lạc sang buổi khác không."""
    pr_low = re.sub(r"\s+", " ", (prompt_text or "").lower())
    es_low = re.sub(r"\s+", " ", (essay_text or "").lower())
    if re.search(
        r"today\s+i\s+write\s+(?:an\s+|this\s+)?(?:email|letter)\s+to\s+(?:talk\s+about|give\s+information|tell\s+you)|"
        r"i'm\s+write\s+spent\s+you|long\s+time\s+no\s+(?:see|meet)\s+you|"
        r"the\s+first\s+one\b[\s\S]{0,120}\bthe\s+second\b|the\s+last\s+question\s+i\s+want\s+to\s+know|"
        r"this\s+is\s+my\s+opinion\s+and\s+what\s+your",
        es_low,
    ):
        return True
    # Đề hỏi về bữa trưa (lunch) nhưng lại kể lể giờ thức dậy buổi sáng / đánh răng rửa mặt
    if is_part1 and "lunch" in pr_low and "wake up" not in pr_low and re.search(r"\b(?:wake|get)\s+up\s+at\b|\bbrush\s+my\s+teeth\b", es_low):
        return True
    return False


def _count_part1_valid_friend_questions(essay_text: str) -> int:
    """
    Đếm số câu hỏi hợp lệ hỏi về thói quen/sở thích của bạn trong Part 1:
    - Loại bỏ các câu hỏi xã giao hoặc câu mời đi chơi không thuộc yêu cầu đề ('How are you?', 'Would you like to my house?', 'OK for you?', 'Is that OK?').
    - Chấp nhận cả câu hỏi đúng chủ đề nhưng sinh viên quên dấu '?' ở cuối câu (VD: 'What time do you wake up. And what your usually does in the morning.').
    """
    cleaned = re.sub(
        r"\b(?:how\s+are\s+you(?:\s+doing)?|how's\s+it\s+going|would\s+you\s+like\s+to\s+(?:go\s+to\s+|come\s+to\s+)?my\s+house|"
        r"is\s+that\s+ok(?:\s+for\s+you)?|ok\s+for\s+you|are\s+you\s+free[^.?!]*|can\s+you\s+come[^.?!]*)\s*\??",
        " ",
        essay_text or "",
        flags=re.IGNORECASE,
    )
    q_marks = cleaned.count("?")
    unpunct_topic_qs = len(
        re.findall(
            r"\b(?:what(?:\s+time|\s+kind|\s+food|\s+sport|\s+room)?|where|when|how(?:\s+often|\s+many)?|who)\s+"
            r"(?:do|does|are|is|can|will)?\s*(?:you|your)\b|"
            r"\blet\s+me\s+know\s+(?:what|where|when|how|who)\b",
            cleaned,
            re.IGNORECASE,
        )
    )
    return max(q_marks, unpunct_topic_qs)


def _calibrate_part_scores(
    prompt_text: str,
    essay_text: str,
    c: int,
    o: int,
    g: int,
    v: int,
    is_part1: bool,
    format_type: str = "BINH_THUONG",
) -> tuple[int, int, int, int, list[str]]:
    """Hiệu chỉnh điểm hậu kỳ theo đúng bộ quy tắc chấm mẫu của Giảng viên IUH (v7 - Gold Standard 2607AVCB005 & 2607AVCB006)."""
    notes = []
    w = count_words(essay_text)
    # 1. Bỏ trống hoặc viết dở dang 1 câu cụt chưa có nội dung (< 12 từ) -> 0 toàn bộ
    if w < 12:
        return 0, 0, 0, 0, ["chưa có nội dung để chấm (<12 từ)"]

    # 2. Chép lại nguyên văn phần lớn đề bài -> 0 toàn bộ
    is_half_off_topic = False
    own_words, copy_ratio = _count_non_copied_words(prompt_text, essay_text)
    if copy_ratio >= 0.55 and own_words < 15:
        notes.append("chép lại nguyên văn phần lớn đề bài (0 điểm toàn bộ)")
        return 0, 0, 0, 0, notes
    elif copy_ratio >= 0.45 and own_words < 25:
        c = min(c, 2)
        o = min(o, 2)
        g = min(g, 2)
        v = min(v, 2)
        is_half_off_topic = True
        notes.append("lặp lại nhiều câu từ đề bài")

    # 3. Nhận diện văn mẫu (mời đi xem phim / hẹn giờ I'm free at... / bận học tiếng Anh thi thứ Hai / đi bảo tàng / hẹn đi Highlands Coffee) chép vào đề khác chủ đề
    pr_low = re.sub(r"\s+", " ", (prompt_text or "").lower())
    es_low = re.sub(r"\s+", " ", (essay_text or "").lower())
    pinfo = _extract_prompt_info(prompt_text, is_part1)
    sender_name_low = (pinfo.get("sender_name") or "").strip().lower()
    greet_name, sign_name = _extract_greeting_and_signoff_names(essay_text)

    if "cinema" not in pr_low and "movie" not in pr_low:
        if re.search(
            r"\b(?:would|wuold)\s+you\s+like\s+(?:to\s+)?(?:come|go|to\s+the\s+weekend)\b[\s\S]{0,80}\bfree\b",
            essay_text or "",
            re.IGNORECASE,
        ):
            c = 0
        elif is_part1 and re.search(r"watch\s+film\s+at\s+cinema\s+with\s+me", es_low):
            c = min(c, 2)
            is_half_off_topic = True
        elif (
            not is_part1
            and re.search(r"busy\s+at[\s\S]{0,100}(?:exam\s+on\s+monday|study\s+at\s+english)", essay_text or "", re.IGNORECASE)
            and ("cinema" in es_low or "watching film" in es_low)
        ):
            c = 0

    # 3-new-A. Part 1 nhưng học thuộc văn mẫu đặt lịch hẹn ("Would you like to [eat/play/listen/go]... I'm free at [giờ]... Let's meet / Is that OK")
    if is_part1 and re.search(r"\b(?:would|wuold)\s+you\s+like\s+to\b", es_low) and re.search(
        r"\b(?:i'm\s+free|i\s+am\s+free|are\s+you\s+free\s+on|what\s+time\s+are\s+you\s+free|let's\s+meet\s+at)\b",
        es_low,
    ):
        c = 0
        notes.append("học thuộc văn mẫu hẹn giờ/mời đi chơi (Would you like to... I'm free...) không khớp đề Part 1")

    # 3-new-B. Part 1 nhưng biến toàn bộ các ý kể về bản thân thành câu hỏi dồn dập cho bạn (VD: Huỳnh Đan Trường P1)
    if is_part1 and re.search(
        r"what\s+time\s+do\s+you\s+wake\s+up[\s\S]{0,120}what\s+do\s+you\s+do\s+in\s+the\s+morning[\s\S]{0,120}what\s+do\s+you\s+have\s+for\s+breakfast",
        es_low,
    ):
        c = 0
        notes.append("biến các gạch đầu dòng kể về bản thân thành câu hỏi ngược lại cho người nhận")

    if not is_part1 and re.search(r"busy\s+at[\s\S]{0,100}study\s+at\s+english", essay_text or "", re.IGNORECASE):
        is_half_off_topic = True
    if not is_part1 and "museum" not in pr_low and re.search(r"visit\s+the\s+museum[\s\S]{0,120}ho\s+chi\s+minh\s+museum", essay_text or "", re.IGNORECASE):
        c = 0
    if not is_part1 and "museum" in pr_low and re.search(r"wellcome\s+to\s+my\s+house|my\s+house\s+is\s+three\s+room", es_low):
        c = 0

    # 3a. Nhận diện đảo ngược vai giao tiếp / học thuộc văn mẫu ở Part 2:
    if not is_part1:
        if "coffee" not in pr_low and re.search(r"highlands?\s+coffee|drink\s+milk\s+tea", es_low):
            c = 0
            notes.append("lạc đề văn mẫu (hẹn đi Highlands Coffee / uống trà sữa không khớp đề)")
        elif (
            "thank you for inviting me to your" in pr_low
            and re.search(r"thank\s+you\s+for\s+your\s+invitation", es_low)
            and re.search(r"do\s+you\s+want\s+me\s+to\s+bring", es_low)
        ):
            c = 0
            notes.append("đảo ngược vai người nhận thư và chép lại câu hỏi của đề")
        # Ký tên cuối thư bằng chính tên người gửi trên đề (VD: đề Sophie gửi mà ký tên Sophie, hoặc chào Hi Alex và ký tên Alex) kèm văn mẫu bận thi / cảm ơn lời mời
        elif sender_name_low and sign_name == sender_name_low and (
            greet_name == sender_name_low
            or re.search(r"busy|exam|thank\s+you\s+for\s+inviting\s+me", es_low)
        ):
            c = 0
            notes.append(f"nhầm vai người viết (ký tên trùng người gửi '{pinfo.get('sender_name')}') và ghép văn mẫu")
        # Đề hỏi kế hoạch tương lai (museum/picnic/concert/party/camping) nhưng lại kể chuyện quá khứ đã đi Vũng Tàu / ăn hải sản
        elif re.search(r"\bi\s+went\s+to\s+vung\s+tau\b|\bwe\s+ate\s+seafood\s+and\s+took\s+many\s+photos\b", es_low):
            c = 0
            notes.append("lạc đề (kể chuyện quá khứ đi Vũng Tàu thay vì trả lời câu hỏi kế hoạch trong đề)")

    # 3b. Nếu Part 1 mà sinh viên tự viết luôn 2 bức thư đối đáp (có >= 2 dòng mở đầu bằng Hi/Hello/Dear/Good...) -> Content = 1
    if is_part1 and c > 1:
        lines = [ln.strip() for ln in (essay_text or "").splitlines() if ln.strip()]
        greeting_lines = [
            ln
            for ln in lines
            if re.match(r"^(hi|hello|dear|good\s+(morning|afternoon|evening))\b", ln, re.IGNORECASE)
            and count_words(ln) <= 5
        ]
        if len(greeting_lines) >= 2:
            c = 1

    # 3b-2. Nếu Part 2 mà sinh viên không phân biệt được ngôi giao tiếp với người nhận thư (xưng hô với chính người nhận thư ở ngôi thứ 3 như "want Emma to bring...", "Emma can't help you prepare...") -> Content = 1 (phạt dây chuyền 1-1-1-1, VD: Đặng Thị Ngọc Hà P2)
    if not is_part1 and re.search(
        r"\bwant\s+emma\s+to\b|\bemma\s+can't\s+help\s+you\b|\bonly\s+emma\s+is\b|\bhope\s+emma\s+not\b",
        es_low,
    ):
        c = 1
        o = min(o, 2)
        g = min(g, 2)
        v = min(v, 2)
        notes.append("sai ngôi giao tiếp (gọi người nhận thư Emma ở ngôi thứ 3 và nhại lại đề)")

    # 3c. Nếu Content = 0 (lạc đề hoàn toàn / học thuộc văn mẫu sai đề) -> Khóa cứng 0 - 1 - 1 - 1
    if c == 0:
        notes.append("lạc đề hoàn toàn (khóa 0-1-1-1)")
        return 0, 1, 1, 1, notes

    # 4. Nếu Content = 1 (không hiểu đề / sai ngôi giao tiếp hoàn toàn) hoặc Content = 2 do "nửa đúng đề nửa học tủ / chép đề":
    # Các cột còn lại bị phạt trừ 1 điểm (sàn tối thiểu là 1 nếu đã đạt >= 1).
    if c == 1 or (c == 2 and (is_half_off_topic or w < 20)):
        o = max(1, min(o - 1, 2)) if o >= 1 else 0
        g = max(1, min(g - 1, 2)) if g >= 1 else 0
        v = max(1, min(v - 1, 2)) if v >= 1 else 0
        if c == 1:
            return c, o, g, v, notes

    # 5. Kiểm tra độ dài quá ngắn (< 30 từ)
    if w < 20:
        c, o, g, v = min(c, 2), min(o, 2), min(g, 2), min(v, 2)
    elif w < 30:
        c, o, g, v = min(c, 3), min(o, 3), min(g, 3), min(v, 3)

    # 6. Đối với Part 1: kiểm tra câu hỏi hợp lệ cho bạn (đã loại bỏ 'How are you?', 'Would you like to my house?', 'OK for you?', và chấp nhận câu hỏi quên dấu '?')
    valid_p1_qs = _count_part1_valid_friend_questions(essay_text) if is_part1 else 0
    if is_part1 and valid_p1_qs == 0:
        c = min(c, 3)

    # 6a. Đối với Part 1 L006 (favourite food / how often / where you like to eat it): nếu lái sang kể bữa sáng (breakfast) và viết câu cụt "This is my opinion and what your favourite food" (VD: Ngô Tuấn Tú P1) -> Content = 3
    if is_part1 and "favourite food" in pr_low and "how often you eat it" in pr_low:
        if "breakfast" in es_low and re.search(r"this\s+is\s+my\s+opinion\s+and\s+what\s+your\s+favourite\s+food", es_low):
            c = min(c, 3)

    # 6b. Đối với Part 1: nếu không đặt câu hỏi trực tiếp cho bạn mà lại kể về bạn ở ngôi thứ 3 ("My friend's favourite sport is... She can play...") -> thiếu câu hỏi trực tiếp (Content = 3)
    if is_part1 and re.search(r"\bmy\s+friend's\s+favourite\s+\w+\s+is\b", es_low) and valid_p1_qs == 0:
        c = min(c, 3)
    elif is_part1 and re.search(
        r"\b(?:where|what|how|who)\s+(?:does|is|can)?\s*(?:my\s+friend|your\s+friend|he|she)\b",
        essay_text or "",
        re.IGNORECASE,
    ):
        c = min(c, 4)
        g = min(g, 4)
    if is_part1 and re.search(r"give\s+information\s+about\s+my\s+city|special\s+about\s+my\s+city", essay_text or "", re.IGNORECASE):
        c = min(c, 3)
        o = min(o, 2)
        g = min(g, 2)

    # 6c. Nếu bài viết dài kèm lan man / câu dẫn sáo rỗng (hoặc quá dài >= 115 từ) -> trừ 1 điểm Content
    is_rambling = _has_rambling_or_meta_filler(prompt_text, essay_text, is_part1)
    if (w >= 78 and is_rambling) or w >= 115:
        if is_part1 and valid_p1_qs == 0:
            c = min(c, 2)
            o = min(o, 3)
        else:
            c = min(c, 4)

    # 7. Thiếu kết thư hoặc thiếu ký tên người viết ở cuối bài (kể cả chỉ ghi 'Your friend.' không tên) -> trừ 1 điểm Content (tối đa 4)
    if _missing_signoff_or_closing(essay_text):
        c = min(c, 4)

    # 7b. Đặt vào vị trí người nhận thư: đề hỏi "What time are we going?" mà chỉ nói "The concert starts at 7p.m." (không nói mấy giờ đi)
    if (
        not is_part1
        and "what time are we going" in pr_low
        and "concert starts at" in es_low
        and not re.search(r"\b(go|going|leave)\b[^.?!]{0,25}\b\d", es_low)
    ):
        c = min(c, 3 if _missing_signoff_or_closing(essay_text) else 4)
        o = min(o, 3)

    # 7c. Đặt vào vị trí người nhận thư ở Part 2: phân hóa lỗi đảo vai chủ-khách (4đ vs 3đ) và kiểm tra thiếu ý câu hỏi
    if not is_part1:
        # Đảo vai nhẹ (4đ): Mở đầu bằng "Would you like to come (with me)... party" nhưng vẫn trả lời đủ 3 câu hỏi với tư cách chủ tiệc (VD: Phan Thị Lan Anh P2, Đặng Thị Như Vàng P2)
        if ("inviting me to your party" in pr_low or "inviting me to your picnic" in pr_low) and re.search(
            r"\bwould\s+you\s+like\s+to\s+come\b", es_low
        ):
            c = min(c, 4)
        # Đảo vai 2 ý (3đ): Mở đầu "Thanks you for inviting" VÀ trả lời "I can bring some food and drinks" khi bạn hỏi "Can I bring some food or drinks?" (VD: Phạm Thị Bình Chi P2)
        if re.search(r"thanks?\s+for\s+inviting\s+me\s+to\s+your", pr_low) and re.search(
            r"thanks?\s+(?:you\s+)?for\s+inviting", es_low
        ) and re.search(r"\bi\s+can\s+bring\s+some\s+food", es_low):
            c = min(c, 3)
        # Nhại lại câu mở đầu của đề ("Thank you for inviting me to the cinema! I'd love to come") kèm ghép câu bận học/thi ("have a lesson... for my exam", VD: Bùi Phước Khoa P2)
        if re.search(r"thank\s+you\s+for\s+inviting\s+me\s+to\s+the\s+\w+\s*!\s*i'd\s+love\s+to\s+come", es_low) and re.search(
            r"\b(?:exam|lesson|busy)\b", es_low
        ):
            c = min(c, 3)
            o = min(o, 2)

        # Lê Nguyễn Tấn Phong (P2): Trả lời rõ 2/3 câu (giờ/rạp hẹn + ăn Phở/Bánh mì) và có chào/kết thư, chỉ lỗi cấu trúc câu tên phim ("Can you watch the movie title happy") -> 4 - 3 - 3 - 3
        if "movie title" in pr_low and re.search(r"can\s+you\s+watch\s+the\s+movie\s+title", es_low):
            c = min(c, 4)
            o = min(o, 3)
            g = min(g, 3)
            v = min(v, 3)
        if "how you want to go to the club" in pr_low and not re.search(
            r"\b(motorbike|motor|bus|car|walk|bike|bicycle|taxi|grab|train|foot|drive|ride)\b", es_low
        ):
            c = min(c, 3 if not re.search(r"\b(bring|food|drink|water|shoes|shirt|ball)\b", es_low) else 4)
            o = min(o, 4)
            g = min(g, 4)
        if "eat something with me" in pr_low and not re.search(
            r"\b(eat|food|restaurant|noodles|dinner|lunch|drink|pizza|chicken|soup|pho|banh|rice|sushi|hotpot|bbq)\b", es_low
        ):
            c = min(c, 4)
            o = min(o, 3)
            v = min(v, 3)
        if "bring anything for the visit" in pr_low and not re.search(
            r"\b(bring|fruit|snack|drink|water|umbrella|food|camera|hat|money|nothing|anything|yourself)\b", es_low
        ):
            c = min(c, 3)
        if "what you want" in pr_low and "for the picnic" in pr_low and re.search(r"let\s+me\s+know\s+what\s+you\s+want", es_low):
            c = min(c, 3)
            o = min(o, 3)
        if "what time should i be there" in pr_low and not re.search(r"\b\d+\s*(?:a\.?m|p\.?m|o'clock|:)|\b(?:morning|afternoon|evening)\b", es_low):
            c = min(c, 4)
            o = min(o, 3)
        if "camping trip" in pr_low and re.search(r"^hi\s+son\b", es_low) and "exam next monday" in es_low:
            c = min(c, 3)
            o = min(o, 2)
            g = min(g, 3)
            v = min(v, 3)

    # 8. Nếu bị lỗi Xuống dòng rời rạc sau mỗi câu đơn lẻ (XUONG_DONG_TUNG_CAU, VD: Lê Hoàng An P2, Phạm Thị Bình Chi P2) -> trừ điểm Organization (tối đa 3)
    if format_type == "XUONG_DONG_TUNG_CAU":
        o = min(o, 3)

    # 8a. Nếu lặp cấu trúc đầu câu "I'm / I am / I ..." từ 4 câu trở lên -> Organization tối đa 4
    # Đặc biệt nếu lặp >= 5 câu đơn bắt đầu bằng "I..." mà không có từ nối chuyển ý (because, so, then, after that, besides...) -> Organization tối đa 3 (VD: Hoàng Kiều Anh P1)
    i_starts_count = len(re.findall(r"(?:^|[.?!]\s+)i(?:'m|\s+am)?\b", essay_text or "", re.IGNORECASE))
    has_connectors = bool(re.search(r"\b(because|so|then|after\s+that|besides|first|second|next|finally|however|moreover)\b", es_low))
    if i_starts_count >= 5 and not has_connectors:
        o = min(o, 3)
    elif i_starts_count >= 4:
        o = min(o, 4)
    if is_part1 and re.search(r"\?\s+i\s+(?:can|usually|often|always|like|love)\b", essay_text or "", re.IGNORECASE):
        o = min(o, 4)

    # 8b. Nếu dùng dấu phẩy ',' thay vì dấu chấm câu '.' nối các mệnh đề độc lập (Comma splice):
    # Trừ điểm Organization xuống tối đa 3 (hoặc 2 nếu nối dấu phẩy tràn lan >= 3 chỗ)
    comma_splices = _count_comma_splices_and_fragments(essay_text or "")
    if comma_splices >= 3:
        o = min(o, 2)
    elif comma_splices >= 1:
        o = min(o, 3)

    # 9. Kiểm soát Vocabulary:
    # - Nếu có >= 2 lỗi chính tả/dùng từ nặng ("tow rooms", "bead room", "My love my house", "Reply soon me", "intiving", "vagetalbal") -> tối đa 2 điểm (VD: Nguyễn Minh Dương P1).
    # - Nếu có >= 1 lỗi chính tả mà không có >= 2 cụm từ tốt bù lại -> tối đa 3 điểm.
    misspellings = _COMMON_MISSPELLINGS.findall(essay_text or "")
    real_misspellings = [
        m for m in misspellings
        if m.lower() not in ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
        or m.islower()
    ]
    good_phrases = _GOOD_VOCAB_PHRASES.findall(essay_text or "")
    if len(real_misspellings) >= 2 and re.search(r"my\s+love\s+my\s+house|reply\s+soon\s+me|intiving|vagetalbal", es_low):
        v = min(v, 2)
    elif len(real_misspellings) >= 1:
        v = min(v, 4)
        if len(good_phrases) < 2:
            v = min(v, 3)

    # 10. Kiểm soát Grammar sát thực tế chấm của Giám khảo:
    # - Có >= 3 lỗi cấu trúc nặng -> tối đa 2 điểm.
    # - Có >= 2 lỗi ngữ pháp (hoặc >= 1 lỗi nặng + >= 1 lỗi khác) -> tối đa 3 điểm (chấm 3 ngay từ đầu thay vì chấm 4 rồi đợi GV hạ xuống 3).
    # - Có 1 lỗi ngữ pháp -> tối đa 4 điểm.
    severe_errs = _SEVERE_GRAMMAR_ERRORS.findall(essay_text or "")
    minor_errs = _COMMON_GRAMMAR_ERRORS.findall(essay_text or "")
    distinct_errs = list(dict.fromkeys([e.strip().lower() for e in (severe_errs + minor_errs) if e.strip()]))
    if len(severe_errs) >= 3:
        g = min(g, 2)
    elif len(distinct_errs) >= 2 or len(severe_errs) >= 2:
        g = min(g, 3)
    elif len(distinct_errs) >= 1:
        g = min(g, 4)

    return c, o, g, v, notes


def _is_generic_or_stale_reason(
    reason_text: str,
    score: int,
    crit_key: str = "",
    has_specific_errors: bool = False,
) -> bool:
    """
    Kiểm tra xem câu nhận xét hiện tại có bị trống, chung chung (không chỉ rõ lỗi cho Giám khảo kiểm chứng),
    hoặc bị lệch với mức điểm thực tế không (ví dụ điểm 3 nhưng vẫn hiện 'Đạt: Trả lời đầy đủ...').
    """
    s = (reason_text or "").strip()
    if not s:
        return True
    generic_prefixes = (
        "Lỗi: Thiếu khoảng 2 ý/câu hỏi",
        "Lỗi: Các câu rời rạc, thiếu liên kết",
        "Lỗi: Ngữ pháp yếu, sai cấu trúc ở nhiều câu",
        "Giới hạn (3/5): Chỉ dùng từ vựng cơ bản",
        "Đạt cấu trúc câu chính, có 1-2 lỗi ngữ pháp nhẹ.",
        "Đạt: Trình tự câu hợp lý, ngắt câu đúng nhưng ít từ nối",
        "Đạt: Trả lời đầy đủ các yêu cầu của đề, có kết thư và ký tên.",
        "Đạt phần lớn yêu cầu, thiếu hoặc lệch nhẹ 1 chi tiết (-1đ).",
        "Đạt: Có cụm từ vựng phù hợp chủ đề, diễn đạt khá tự nhiên.",
        "Lỗi: Có một số lỗi chia động từ / cấu trúc câu hỏi cơ bản.",
        "Ngữ pháp trung bình (3đ): Còn một số lỗi chia động từ, thiếu trợ động từ trong câu hỏi hoặc lỗi cấu trúc câu ghép.",
    )
    if any(s.startswith(gp) for gp in generic_prefixes):
        return True
    # Nếu trong lý do có ghi "(Xđ)" hoặc "(X/5)" hoặc "mức X điểm" mà X != score -> đã bị lệch điểm
    m_score = re.search(r"(?:\((\d)\s*(?:đ|/5)\)|mức\s+(\d)\s*điểm|tối\s+đa\s+(\d)\s*điểm)", s, re.IGNORECASE)
    if m_score:
        matched_num = next((int(g) for g in m_score.groups() if g is not None), None)
        if matched_num is not None and matched_num != int(score):
            return True
    if score == 0 and "0đ" not in s and "0 điểm" not in s and "Bỏ trống" not in s:
        return True
    # Khắc phục triệt để lỗi điểm <= 3 nhưng vẫn giữ câu khen của mức 4-5 điểm ("Đạt: Trả lời đầy đủ...", "Mạch văn logic...", "Đạt: Mạch lạc...")
    if score <= 3:
        praise_prefixes_for_4_5 = (
            "Đạt:",
            "Đạt ",
            "Trả lời đầy đủ",
            "Trả lời trọn vẹn",
            "Mạch lạc",
            "Mạch văn logic",
            "Cấu trúc đoạn văn logic",
            "Bố cục tốt",
            "Bố cục khá",
            "Ngữ pháp rất tốt",
            "Ngữ pháp tốt",
            "Ngữ pháp khá",
            "Từ vựng tốt",
            "Từ vựng khá",
            "Từ vựng phong phú",
            "Dùng từ vựng tự nhiên",
        )
        if any(s.startswith(pp) for pp in praise_prefixes_for_4_5):
            return True
    # Nếu ở cột Grammar (score <= 4) mà câu nhận xét của AI không trích dẫn cụm từ lỗi nào trong dấu ngoặc kép, trong khi bộ dò đã bắt được lỗi cụ thể -> cập nhật để trích dẫn rõ cho Giám khảo kiểm chứng
    if crit_key == "g" and score <= 4 and has_specific_errors and '"' not in s and "'" not in s:
        return True
    if score >= 4 and (
        s.startswith("Lỗi:")
        or s.startswith("Lạc đề")
        or s.startswith("Khóa điểm")
        or s.startswith("Thiếu toàn bộ")
        or s.startswith("Đạt một phần (3đ)")
        or s.startswith("Ngữ pháp trung bình (3đ)")
        or s.startswith("Từ vựng cơ bản (3đ)")
        or s.startswith("Từ vựng cơ bản & có lỗi chính tả (3đ)")
    ):
        return True
    return False


def _synthesize_part_cogv_reasons(
    prompt_text: str,
    essay_text: str,
    c: int,
    o: int,
    g: int,
    v: int,
    rc: str,
    ro: str,
    rg: str,
    rv: str,
    is_part1: bool,
    format_type: str = "BINH_THUONG",
    notes: list[str] | None = None,
    force_rebuild: bool = False,
) -> tuple[str, str, str, str]:
    """
    Tổng hợp giải thích chi tiết dành cho Giám khảo kiểm chứng (Examiner Verification Format):
    Chỉ rõ đạt/thiếu ý nào của đề và trích dẫn nguyên văn cụm từ sai ("...") tương ứng với mức điểm C, O, G, V.
    """
    notes = notes or []
    w = count_words(essay_text)
    if w == 0:
        msg = "Bỏ trống (0 từ) -> 0 điểm."
        return msg, msg, msg, msg
    if w < 12 or any("<12 từ" in n for n in notes):
        msg = f"Bài viết quá cụt ({w} từ < 12 từ), chưa đủ nội dung chấm -> 0 điểm."
        return msg, msg, msg, msg
    own_words, copy_ratio = _count_non_copied_words(prompt_text, essay_text)
    if (copy_ratio >= 0.55 and own_words < 15) or any("chép lại nguyên văn" in n for n in notes):
        msg = f"Chép lại nguyên văn phần lớn đề bài ({int(copy_ratio * 100)}% trùng đề, chỉ tự viết {own_words} từ) -> 0 điểm toàn bộ."
        return msg, msg, msg, msg

    pr_low = re.sub(r"\s+", " ", (prompt_text or "").lower())
    es_low = re.sub(r"\s+", " ", (essay_text or "").lower())
    pinfo = _extract_prompt_info(prompt_text, is_part1)
    info_bullets = pinfo["info_bullets"]
    ask_bullets = pinfo["ask_bullets"]
    p2_questions = pinfo["p2_questions"]
    sender_name = pinfo["sender_name"]
    greet_name, sign_name = _extract_greeting_and_signoff_names(essay_text)

    misspellings = _COMMON_MISSPELLINGS.findall(essay_text or "")
    real_mis = list(
        dict.fromkeys(
            m for m in misspellings
            if m.lower() not in ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
            or m.islower()
        )
    )
    good_phrases = list(dict.fromkeys(_GOOD_VOCAB_PHRASES.findall(essay_text or "")))
    severe_errs = _SEVERE_GRAMMAR_ERRORS.findall(essay_text or "")
    minor_errs = _COMMON_GRAMMAR_ERRORS.findall(essay_text or "")
    found_g = list(dict.fromkeys([e.strip() for e in (severe_errs + minor_errs) if e.strip()]))
    comma_splices = _count_comma_splices_and_fragments(essay_text or "")
    valid_p1_qs = _count_part1_valid_friend_questions(essay_text) if is_part1 else 0

    # Trường hợp Content = 0 (lạc đề hoàn toàn / học thuộc văn mẫu sai đề)
    if c == 0:
        if any("tự viết thư hồi đáp cho Part 1" in n for n in notes):
            rc_out = (
                f"Lạc đề hoàn toàn (0đ): Đề Part 2 yêu cầu trả lời thư của {sender_name or 'bạn trên đề'}, "
                f"nhưng sinh viên lại tự đóng vai người nhận ở Part 1 để viết thư hồi đáp cho chính bài Part 1 của mình (chào 'Hi {greet_name.title()}' và ký tên '{sign_name.title()}')."
            )
        elif any("biến các gạch đầu dòng kể về bản thân" in n for n in notes):
            rc_out = (
                "Lạc đề hoàn toàn (0đ): Đề yêu cầu kể thông tin bản thân (wake up, morning activities, breakfast), "
                "nhưng sinh viên lại biến toàn bộ các ý đó thành câu hỏi dồn dập cho bạn ('What time do you wake up? What do you do in the morning? What do you have for breakfast?')."
            )
        elif any("học thuộc văn mẫu hẹn giờ/mời đi chơi" in n for n in notes):
            rc_out = (
                "Lạc đề hoàn toàn (0đ): Đề Part 1 yêu cầu kể thói quen/sở thích và đặt câu hỏi cho bạn, "
                "nhưng bài làm lại chép văn mẫu đặt lịch hẹn giờ ('Would you like to... I'm free at... Let's meet...')."
            )
        elif any("nhầm vai người viết (ký tên trùng người gửi" in n for n in notes):
            rc_out = (
                f"Lạc đề hoàn toàn (0đ): Nhầm vai người viết thư (ký tên cuối thư là '{sign_name.title()}' trùng với tên người gửi trên đề) "
                "và ghép văn mẫu bận học/thi ('busy / have an exam') không trả lời câu hỏi của đề."
            )
        elif any("kể chuyện quá khứ đi Vũng Tàu" in n for n in notes):
            rc_out = (
                "Lạc đề hoàn toàn (0đ): Đề hỏi kế hoạch cho buổi hẹn sắp tới nhưng bài làm lại chép văn mẫu kể chuyện quá khứ đi Vũng Tàu ('I went to Vung Tau... ate seafood')."
            )
        elif "coffee" not in pr_low and re.search(r"highlands?\s+coffee|drink\s+milk\s+tea", es_low):
            rc_out = (
                "Lạc đề hoàn toàn (0đ): Đề bài bạn là chủ bữa tiệc trả lời thư của Emma, nhưng bài làm lại chép văn mẫu "
                "cảm ơn lời mời ('Thank you for your invitation'), hẹn đi Highlands Coffee uống trà sữa và hỏi ngược lại câu hỏi của đề."
            )
        elif "cinema" not in pr_low and ("cinema" in es_low or "movie" in es_low or "film" in es_low):
            rc_out = (
                "Lạc đề hoàn toàn (0đ): Bài làm học thuộc văn mẫu hẹn đi xem phim / bận học tiếng Anh thi thứ Hai, "
                "không liên quan đến chủ đề và câu hỏi trên đề bài."
            )
        elif not is_part1 and "museum" not in pr_low and "museum" in es_low:
            rc_out = "Lạc đề hoàn toàn (0đ): Bài làm chép văn mẫu đi thăm bảo tàng (Ho Chi Minh Museum), lệch hoàn toàn với đề bài."
        elif not is_part1 and "museum" in pr_low and "my house" in es_low:
            rc_out = "Lạc đề hoàn toàn (0đ): Đề hỏi về chuyến đi bảo tàng nhưng bài làm lại tả ngôi nhà 3 phòng ('welcome to my house')."
        elif not force_rebuild and not _is_generic_or_stale_reason(rc, 0, "c"):
            rc_out = rc.strip()
        else:
            rc_out = "Lạc đề hoàn toàn (0đ): Nội dung bài viết học tủ văn mẫu hoặc lệch hẳn yêu cầu giao tiếp của đề bài."

        ro_lock = f"Khóa điểm ({o}đ): Do Content = 0 (lạc đề/học tủ văn mẫu) nên Organization khóa cứng theo quy tắc 0 - 1 - 1 - 1."
        rg_lock = (
            f"Khóa điểm ({g}đ): Do Content = 0 (lạc đề) nên Grammar khóa ở mức {g}đ"
            + (f" (trong bài có lỗi: {', '.join(f'\"{x}\"' for x in found_g[:3])})." if found_g else ".")
        )
        rv_lock = (
            f"Khóa điểm ({v}đ): Do Content = 0 (lạc đề) nên Vocabulary khóa ở mức {v}đ"
            + (f" (kèm lỗi chính tả: {', '.join(f'\"{m}\"' for m in real_mis[:3])})." if real_mis else ".")
        )
        return rc_out, ro_lock, rg_lock, rv_lock

    # 1. Tổng hợp chi tiết cho Content (c = 1..5)
    if not force_rebuild and not _is_generic_or_stale_reason(rc, c, "c"):
        rc_out = rc.strip()
    else:
        if c == 5:
            if is_part1:
                req_summary = f"đủ {len(info_bullets) or 3} ý thông tin ({'; '.join(info_bullets[:3])}) & {len(ask_bullets) or 2} câu hỏi cho bạn"
            else:
                req_summary = f"trả lời trọn vẹn {len(p2_questions) or 3} câu hỏi của {sender_name or 'bạn'}"
            rc_out = f"Đạt trọn vẹn (5đ): Đáp ứng {req_summary}, đúng vai người viết thư ({w} từ), có lời chào và kết thư/ký tên đầy đủ."
        elif c == 4:
            if not is_part1 and "movie title" in pr_low and re.search(r"can\s+you\s+watch\s+the\s+movie\s+title", es_low):
                rc_out = (
                    "Đạt 2/3 câu hỏi + đủ chào/kết thư (4đ): Trả lời rõ giờ/nơi hẹn (9h tối tại Galaxy Tân Sơn Nhất) và món ăn "
                    "(Phở/Bánh mì), nhưng câu tên phim viết sai cấu trúc (\"Can you watch the movie title happy\" làm người nhận khó hiểu, -1đ)."
                )
            elif not is_part1 and ("inviting me to your party" in pr_low or "inviting me to your picnic" in pr_low) and re.search(r"\bwould\s+you\s+like\s+to\s+come\b", es_low):
                rc_out = (
                    "Lệch nhẹ câu mở đầu (4đ): Trả lời đủ 3 câu hỏi của bạn nhưng câu mở đầu lại viết lời mời mới "
                    "(\"Would you like to come...\") trong khi bạn đã nhận lời mời trên đề (-1đ)."
                )
            elif _missing_signoff_or_closing(essay_text):
                tail_preview = (essay_text or "").strip().splitlines()[-1].strip() if (essay_text or "").strip() else ""
                rc_out = (
                    f"Đạt đủ ý chính nhưng thiếu kết thư/ký tên (4đ): Đã trả lời các yêu cầu của đề ({w} từ) "
                    f"nhưng dòng cuối thư (\"{tail_preview[-35:]}\") thiếu câu kết hoặc thiếu tên người viết (-1đ)."
                )
            elif _has_rambling_or_meta_filler(prompt_text, essay_text, is_part1) or w >= 85:
                rc_out = (
                    f"Đạt đủ ý nhưng viết dài lan man (4đ): Trả lời đủ câu hỏi của đề nhưng bài dài {w} từ "
                    "(chêm câu dẫn nhập sáo rỗng hoặc kể lể chi tiết ngoài trọng tâm đề ~50 từ, -1đ)."
                )
            elif is_part1 and re.search(r"\b(?:where|what|how|who)\s+(?:does|is|can)?\s*(?:my\s+friend|your\s+friend|he|she)\b", essay_text or "", re.IGNORECASE):
                rc_out = "Sai đại từ xưng hô khi đặt câu hỏi (4đ): Đủ các ý nhưng bê nguyên cụm \"your friend / my friend / he / she\" trên đề vào câu hỏi trực tiếp cho bạn (-1đ)."
            elif not is_part1 and "eat something with me" in pr_low and not re.search(r"\b(eat|food|restaurant|noodles|dinner|lunch|drink|pizza|chicken|soup|pho|banh|rice|sushi|hotpot|bbq)\b", es_low):
                rc_out = "Thiếu 1 ý trả lời người nhận (4đ): Đã nêu giờ/nơi hẹn và tên phim nhưng quên trả lời câu hỏi \"Would you like to eat something with me before the movie?\" (-1đ)."
            elif not is_part1 and "what time should i be there" in pr_low and not re.search(r"\b\d+\s*(?:a\.?m|p\.?m|o'clock|:)|\b(?:morning|afternoon|evening)\b", es_low):
                rc_out = "Thiếu 1 chi tiết trả lời (4đ): Đã nêu địa điểm và đồ chuẩn bị nhưng chưa trả lời giờ cụ thể cho câu hỏi \"What time should I be there?\" (-1đ)."
            else:
                reqs = p2_questions if not is_part1 else (info_bullets[:2] + [f"hỏi {x}" for x in ask_bullets[:2]])
                req_str = "; ".join(reqs[:3]) if reqs else "các yêu cầu của đề"
                rc_out = f"Đạt phần lớn yêu cầu (4đ): Trả lời tốt các ý chính ({req_str}), có mở/kết thư nhưng còn 1 chi tiết diễn đạt chưa rõ cho người nhận (-1đ)."
        elif c == 3:
            if is_part1 and valid_p1_qs == 0:
                ask_str = "; ".join(ask_bullets[:2]) if ask_bullets else "2 câu hỏi cho bạn ở phần Remember to ask"
                if re.search(r"\bmy\s+friend's\s+favourite\b", es_low):
                    rc_out = (
                        f"Không đặt câu hỏi trực tiếp cho bạn (3đ): Đã kể thông tin bản thân nhưng phần Remember to ask ({ask_str}) "
                        "lại viết thành câu kể ngôi thứ 3 (\"My friend's favourite sport is... She can play...\") thay vì hỏi người nhận (-2đ)."
                    )
                else:
                    rc_out = (
                        f"Thiếu toàn bộ phần câu hỏi chủ đề cho bạn (3đ): Đã kể thông tin bản thân nhưng quên đặt 2 câu hỏi cho người nhận "
                        f"theo yêu cầu ({ask_str}; các câu mời như \"Would you like to my house?\" / \"OK for you?\" không tính là câu hỏi chủ đề, -2đ)."
                    )
            elif is_part1 and "favourite food" in pr_low and "breakfast" in es_low and re.search(r"this\s+is\s+my\s+opinion\s+and\s+what\s+your\s+favourite\s+food", es_low):
                rc_out = (
                    "Lệch trọng tâm & câu hỏi cụt nghĩa (3đ): Đề hỏi món ăn yêu thích và tần suất ăn (how often you eat it), "
                    "nhưng bài lái sang kể bữa sáng (\"eat meat and egg for my breakfast\") và đặt câu hỏi bị cụt cấu trúc (\"This is my opinion and what your favourite food.\", -2đ)."
                )
            elif not is_part1 and re.search(r"thanks?\s+(?:you\s+)?for\s+inviting", es_low) and re.search(r"\bi\s+can\s+bring\s+some\s+food", es_low):
                rc_out = (
                    "Đảo ngược vai chủ - khách ở 2 ý (3đ): Đề bài bạn là chủ buổi picnic trả lời thư của Ben, nhưng lại mở đầu "
                    "\"Thanks you for inviting\" và tự nhận \"I can bring some food and drinks\" khi Ben hỏi \"Can I bring some food or drinks?\" (-2đ)."
                )
            elif not is_part1 and re.search(r"thank\s+you\s+for\s+inviting\s+me\s+to\s+the\s+\w+\s*!\s*i'd\s+love\s+to\s+come", es_low):
                rc_out = (
                    "Chép lại câu mở đầu của đề & ghép văn mẫu bận học (3đ): Lặp nguyên câu \"Thank you for inviting me to the cinema! I'd love to come\" "
                    "của người gửi và chêm câu bận thi (\"I have a lesson... for my exam\", -2đ)."
                )
            elif not is_part1 and "camping trip" in pr_low and re.search(r"^hi\s+son\b", es_low) and "exam next monday" in es_low:
                rc_out = (
                    "Lệch ngữ cảnh người nhận (3đ): Chào sai tên người nhận (\"Hi Son\"), ghép câu văn mẫu \"exam next Monday\" "
                    "và chưa trả lời rõ phương tiện di chuyển cho chuyến cắm trại (-2đ)."
                )
            elif not is_part1 and "what time are we going" in pr_low and "concert starts at" in es_low:
                rc_out = (
                    "Trả lời lệch trọng tâm người nhận (3đ): Bạn hỏi \"What time are we going?\" (mấy giờ đi) nhưng chỉ nói giờ concert bắt đầu "
                    "(\"starts at 7p.m.\") mà không chốt giờ hẹn đi, kèm thiếu kết thư/ký tên (-2đ)."
                )
            elif not is_part1 and "for the picnic" in pr_low and re.search(r"let\s+me\s+know\s+what\s+you\s+want", es_low):
                rc_out = "Nhại lại câu hỏi của đề (3đ): Bạn hỏi cần chuẩn bị gì cho buổi picnic nhưng lại chép ngược lại câu \"let me know what you want to prepare\" (-2đ)."
            elif not is_part1 and "bring anything for the visit" in pr_low and not re.search(r"\b(bring|fruit|snack|drink|water|umbrella|food|camera|hat|money|nothing|anything|yourself)\b", es_low):
                rc_out = "Thiếu ý trả lời (3đ): Đã trả lời giờ đi và phương tiện nhưng bỏ sót câu hỏi \"Should I bring anything for the visit?\" của bạn (-2đ)."
            else:
                reqs = p2_questions if not is_part1 else (info_bullets[:2] + [f"hỏi {x}" for x in ask_bullets[:2]])
                req_str = "; ".join(reqs[:3]) if reqs else "các ý trên đề"
                sign_note = " Kèm thiếu kết thư/ký tên." if _missing_signoff_or_closing(essay_text) else ""
                rc_out = f"Đạt một phần (3đ): Trong các yêu cầu ({req_str}), bài làm còn thiếu hoặc trả lời lệch khoảng 2 ý từ góc nhìn người nhận (-2đ).{sign_note}"
        elif c == 2:
            if w >= 85 and is_part1 and valid_p1_qs == 0:
                rc_out = f"Thiếu câu hỏi & viết lan man (2đ): Quên hoàn toàn phần đặt 2 câu hỏi cho bạn và viết dài lan man ({w} từ, -3đ)."
            elif copy_ratio >= 0.40:
                rc_out = f"Nửa đúng đề nửa chép đề (2đ): Lặp lại nhiều cụm từ nguyên văn trên đề (chỉ tự viết {own_words} từ), thiếu ý triển khai thực tế (-3đ)."
            else:
                rc_out = f"Thiếu nhiều ý trọng tâm (2đ): Chỉ trả lời được 1 ý nhỏ trong đề ({w} từ), các câu còn lại sơ sài hoặc lệch câu hỏi của người nhận (-3đ)."
        else:  # c == 1
            if not is_part1 and re.search(r"\bwant\s+emma\s+to\b|\bemma\s+can't\s+help\s+you\b|\bonly\s+emma\s+is\b", es_low):
                rc_out = (
                    "Sai ngôi giao tiếp nghiêm trọng (1đ): Viết thư trả lời Emma nhưng không có lời chào mở đầu, "
                    "xưng hô với chính người nhận ở ngôi thứ 3 (\"want Emma to bring...\", \"Emma can't help you prepare...\") và sao chép đề."
                )
            elif is_part1:
                rc_out = "Sai thể thức/yêu cầu (1đ): Không hiểu yêu cầu đề bài (viết thành 2 thư đối đáp hoặc chỉ chạm 1 ý rất nhỏ)."
            else:
                rc_out = "Không đạt mục đích giao tiếp (1đ): Người nhận không hiểu được thông điệp trả lời, sai ngữ cảnh xưng hô."

    # 2. Tổng hợp chi tiết cho Organization (o = 1..5)
    if not force_rebuild and not _is_generic_or_stale_reason(ro, o, "o"):
        ro_out = ro.strip()
    else:
        i_starts = len(re.findall(r"(?:^|[.?!]\s+)i(?:'m|\s+am)?\b", essay_text or "", re.IGNORECASE))
        if o == 5:
            ro_out = "Bố cục tốt (5đ): Tổ chức đoạn văn email mạch lạc, đầy đủ mở - thân - kết thư, dùng từ nối và ngắt câu chuẩn xác."
        elif o == 4:
            if i_starts >= 4:
                ro_out = f"Bố cục khá (4đ): Trình tự các ý hợp lý, đúng chuẩn email nhưng lặp cấu trúc mở đầu câu bằng 'I / I am / I'm' ({i_starts} câu)."
            else:
                ro_out = "Bố cục khá (4đ): Đầy đủ mở - thân - kết thư, trình tự trả lời hợp lý nhưng từ nối chuyển ý còn đơn giản."
        elif o == 3:
            if format_type == "XUONG_DONG_TUNG_CAU":
                ro_out = "Lỗi trình bày đoạn văn (3đ): Xuống dòng rời rạc sau mỗi câu đơn lẻ thay vì viết thành đoạn văn email liền mạch."
            elif comma_splices >= 1:
                ro_out = f"Lỗi ngắt câu & liên kết (3đ): Dùng dấu phẩy ',' thay dấu chấm '.' nối mệnh đề độc lập ({comma_splices} chỗ, comma splice) và ít từ nối."
            elif i_starts >= 5:
                ro_out = f"Bố cục đơn điệu & thiếu liên kết (3đ): Lặp cấu trúc mở đầu câu bằng 'I / I am / I'm' liên tiếp ({i_starts} câu đơn rời rạc) mà không có từ nối chuyển ý."
            else:
                ro_out = "Bố cục trung bình (3đ): Có khung thư cơ bản nhưng cách sắp xếp và chuyển ý giữa các câu trả lời còn rời rạc, đột ngột."
        elif o == 2:
            if format_type == "XUONG_DONG_TUNG_CAU":
                ro_out = "Lỗi trình bày đoạn văn nặng (2đ): Sinh viên xuống dòng rời rạc sau mỗi câu đơn lẻ làm vỡ hoàn toàn bố cục đoạn văn email."
            elif comma_splices >= 2:
                ro_out = f"Lỗi ngắt câu nặng (2đ): Dùng dấu phẩy ',' nối mệnh đề độc lập tràn lan ({comma_splices} chỗ) làm đứt gãy cấu trúc đoạn văn."
            else:
                ro_out = "Bố cục yếu (2đ): Các câu lắp ghép rời rạc, thiếu mạch liên kết tự nhiên giữa các ý trong thư."
        else:
            ro_out = f"Bố cục rất yếu ({o}đ): Thiếu cấu trúc bức thư (không có mở/kết chuẩn), các cụm từ chắp vá rời rạc (phạt dây chuyền từ Content)."

    # 3. Tổng hợp chi tiết cho Grammar (g = 1..5)
    extra_g = []
    if re.search(r"[.?!]\s*would\s+like\s+to\b", essay_text or "", re.IGNORECASE):
        extra_g.append("thiếu chủ ngữ 'I' trước 'would like to'")
    if re.search(r"\bi\s+(?:want|like|need|hope)\s+to?\s*\w+\s+\w+\s+not\b", essay_text or "", re.IGNORECASE):
        extra_g.append("sai cấu trúc phủ định")
    all_g_samples = [f'"{x}"' for x in found_g[:4]] + extra_g[:2]
    g_sample_str = ", ".join(all_g_samples[:4])

    if not force_rebuild and not _is_generic_or_stale_reason(rg, g, "g", has_specific_errors=bool(all_g_samples)):
        rg_out = rg.strip()
    else:
        if g == 5:
            rg_out = "Ngữ pháp tốt (5đ): Kiểm soát tốt cấu trúc câu đơn và câu hỏi, chia động từ và dùng giới từ chuẩn xác."
        elif g == 4:
            if g_sample_str:
                rg_out = f"Ngữ pháp khá (4đ): Cấu trúc câu chính vững, chỉ có lỗi nhẹ ở {g_sample_str}."
            else:
                rg_out = "Ngữ pháp khá (4đ): Các câu đơn và câu hỏi đúng cấu trúc chính, chỉ có 1–2 lỗi nhỏ về mạo từ/giới từ."
        elif g == 3:
            if g_sample_str:
                rg_out = f"Ngữ pháp trung bình (3đ): Đạt câu đơn cơ bản nhưng sai cấu trúc ở {g_sample_str}."
            else:
                rg_out = "Ngữ pháp trung bình (3đ): Còn một số lỗi chia động từ, thiếu trợ động từ trong câu hỏi hoặc lỗi cấu trúc câu ghép."
        elif g == 2:
            if g_sample_str:
                rg_out = f"Nhiều lỗi ngữ pháp nặng (2đ): Sai cấu trúc câu ở {g_sample_str}, làm câu văn tối nghĩa."
            else:
                rg_out = "Nhiều lỗi ngữ pháp nặng (2đ): Sai cấu trúc động từ và trật tự từ ở nhiều câu trong bài."
        else:
            if g_sample_str:
                rg_out = f"Ngữ pháp rất yếu ({g}đ): Sai cấu trúc câu nghiêm trọng ({g_sample_str}) và bị giới hạn theo điểm Content."
            else:
                rg_out = f"Ngữ pháp rất yếu ({g}đ): Hầu hết các câu đều sai cấu trúc ngữ pháp cơ bản (hoặc phạt dây chuyền từ Content)."

    # 4. Tổng hợp chi tiết cho Vocabulary (v = 1..5)
    if not force_rebuild and not _is_generic_or_stale_reason(rv, v, "v"):
        rv_out = rv.strip()
    else:
        if v == 5:
            gp_note = f" (như {', '.join(f'\"{p}\"' for p in good_phrases[:3])})" if good_phrases else ""
            rv_out = f"Từ vựng tốt (5đ): Dùng từ vựng phong phú, tự nhiên đúng chủ đề{gp_note}, không mắc lỗi chính tả."
        elif v == 4:
            if real_mis and good_phrases:
                rv_out = (
                    f"Từ vựng khá (4đ): Có cụm từ tốt ({', '.join(f'\"{p}\"' for p in good_phrases[:2])}) "
                    f"bù trừ cho lỗi chính tả nhỏ ({', '.join(f'\"{m}\"' for m in real_mis[:2])})."
                )
            elif good_phrases:
                rv_out = f"Từ vựng khá (4đ): Từ vựng phù hợp ngữ cảnh giao tiếp, có cụm từ tốt ({', '.join(f'\"{p}\"' for p in good_phrases[:3])})."
            else:
                rv_out = "Từ vựng khá (4đ): Dùng từ vựng đúng chủ đề, diễn đạt rõ nghĩa và không sai chính tả."
        elif v == 3:
            if real_mis:
                rv_out = f"Từ vựng cơ bản & có lỗi chính tả (3đ): Sai chính tả/dùng từ ở {', '.join(f'\"{m}\"' for m in real_mis[:3])}; các từ còn lại ở mức cơ bản."
            else:
                rv_out = "Từ vựng cơ bản (3đ): Chủ yếu dùng các từ đơn giản quen thuộc hoặc lấy lại từ khóa từ đề bài, chưa có cụm từ mở rộng."
        elif v == 2:
            if real_mis:
                rv_out = f"Từ vựng yếu (2đ): Lặp từ trên đề, dùng từ gượng ép và sai chính tả ({', '.join(f'\"{m}\"' for m in real_mis[:3])})."
            else:
                rv_out = "Từ vựng yếu (2đ): Vốn từ hạn chế, lặp lại nguyên mẫu các từ trên đề hoặc dùng sai ngữ cảnh."
        else:
            rv_out = f"Từ vựng rất yếu ({v}đ): Vốn từ quá sơ sài, sao chép từ đề bài hoặc bị phạt dây chuyền do lỗi Content."

    return rc_out, ro_out, rg_out, rv_out


def ensure_cogv_reasons(sv_data: dict, kq: KetQuaCham) -> KetQuaCham:
    """Đảm bảo đối tượng KetQuaCham có đầy đủ 8 trường lý do C-O-G-V và chuỗi nhan_xet tổng hợp."""
    p1 = sv_data.get("p1") or {}
    p2 = sv_data.get("p2") or {}
    pr1 = clean_prompt_text(p1.get("prompt") or "")
    pr2 = clean_prompt_text(p2.get("prompt") or "")
    e1 = (p1.get("essay") or "").strip()
    e2 = (p2.get("essay") or "").strip()
    ft1 = (p1.get("format_type") or "BINH_THUONG").strip()
    ft2 = (p2.get("format_type") or "BINH_THUONG").strip()

    kq.p1_c_reason, kq.p1_o_reason, kq.p1_g_reason, kq.p1_v_reason = _synthesize_part_cogv_reasons(
        pr1, e1,
        kq.p1_content, kq.p1_org, kq.p1_grammar, kq.p1_vocab,
        kq.p1_c_reason, kq.p1_o_reason, kq.p1_g_reason, kq.p1_v_reason,
        is_part1=True, format_type=ft1,
    )
    kq.p2_c_reason, kq.p2_o_reason, kq.p2_g_reason, kq.p2_v_reason = _synthesize_part_cogv_reasons(
        pr2, e2,
        kq.p2_content, kq.p2_org, kq.p2_grammar, kq.p2_vocab,
        kq.p2_c_reason, kq.p2_o_reason, kq.p2_g_reason, kq.p2_v_reason,
        is_part1=False, format_type=ft2,
    )
    if not kq.nhan_xet.strip():
        w1, w2 = count_words(e1), count_words(e2)
        kq.nhan_xet = (
            f"P1 ({w1} từ): C={kq.p1_content} ({kq.p1_c_reason}); O={kq.p1_org} ({kq.p1_o_reason}); "
            f"G={kq.p1_grammar} ({kq.p1_g_reason}); V={kq.p1_vocab} ({kq.p1_v_reason}). | "
            f"P2 ({w2} từ): C={kq.p2_content} ({kq.p2_c_reason}); O={kq.p2_org} ({kq.p2_o_reason}); "
            f"G={kq.p2_grammar} ({kq.p2_g_reason}); V={kq.p2_vocab} ({kq.p2_v_reason})."
        )
    return kq


def _calibrate_scores(sv_data: dict, kq: KetQuaCham) -> KetQuaCham:
    p1 = sv_data.get("p1") or {}
    p2 = sv_data.get("p2") or {}
    pr1 = clean_prompt_text(p1.get("prompt") or "")
    pr2 = clean_prompt_text(p2.get("prompt") or "")
    e1 = (p1.get("essay") or "").strip()
    e2 = (p2.get("essay") or "").strip()
    ft1 = (p1.get("format_type") or "BINH_THUONG").strip()
    ft2 = (p2.get("format_type") or "BINH_THUONG").strip()

    c1, o1, g1, v1, n1 = _calibrate_part_scores(
        pr1, e1, kq.p1_content, kq.p1_org, kq.p1_grammar, kq.p1_vocab, is_part1=True, format_type=ft1
    )
    c2, o2, g2, v2, n2 = _calibrate_part_scores(
        pr2, e2, kq.p2_content, kq.p2_org, kq.p2_grammar, kq.p2_vocab, is_part1=False, format_type=ft2
    )

    # Kiểm tra chéo giữa Part 1 và Part 2:
    # Nếu ở Part 2, sinh viên phớt lờ người gửi trên đề P2 (VD: Ben, Mark, Sophie...) và lại đóng vai người nhận của Part 1
    # để tự viết thư trả lời cho chính bài Part 1 của mình (VD: P1 chào "Hi Hung" ký "Truong", P2 chào "Hi Truong" ký "Hung";
    # hoặc P1 chào "Hi Vy" ký "Toan", P2 chào "Hi Toan" ký "Vy", như Huỳnh Đan Trường 24684671 & Huỳnh Thảo Vy 24694491) -> Khóa P2 = 0 - 1 - 1 - 1
    if count_words(e2) >= 12:
        p1_greet, p1_sign = _extract_greeting_and_signoff_names(e1)
        p2_greet, p2_sign = _extract_greeting_and_signoff_names(e2)
        p2_sender = (_extract_prompt_info(pr2, is_part1=False).get("sender_name") or "").strip().lower()
        if (
            p1_greet
            and p1_sign
            and p2_greet == p1_sign
            and p2_sign == p1_greet
            and p2_greet != p2_sender
        ):
            c2, o2, g2, v2 = 0, 1, 1, 1
            n2.append("lạc đề hoàn toàn: tự viết thư hồi đáp cho Part 1 của chính mình thay vì trả lời đề Part 2 (khóa 0-1-1-1)")

    kq.p1_content, kq.p1_org, kq.p1_grammar, kq.p1_vocab = c1, o1, g1, v1
    kq.p2_content, kq.p2_org, kq.p2_grammar, kq.p2_vocab = c2, o2, g2, v2

    kq.p1_c_reason, kq.p1_o_reason, kq.p1_g_reason, kq.p1_v_reason = _synthesize_part_cogv_reasons(
        pr1, e1, c1, o1, g1, v1,
        kq.p1_c_reason, kq.p1_o_reason, kq.p1_g_reason, kq.p1_v_reason,
        is_part1=True, format_type=ft1, notes=n1,
    )
    kq.p2_c_reason, kq.p2_o_reason, kq.p2_g_reason, kq.p2_v_reason = _synthesize_part_cogv_reasons(
        pr2, e2, c2, o2, g2, v2,
        kq.p2_c_reason, kq.p2_o_reason, kq.p2_g_reason, kq.p2_v_reason,
        is_part1=False, format_type=ft2, notes=n2,
    )

    extra_warnings = []
    if n1:
        extra_warnings.append(f"P1 {', '.join(n1)}")
    if n2:
        extra_warnings.append(f"P2 {', '.join(n2)}")
    if extra_warnings:
        prefix = "⚠️ " + "; ".join(extra_warnings) + ". "
        if prefix.strip() not in kq.nhan_xet:
            kq.nhan_xet = prefix + kq.nhan_xet
    ensure_cogv_reasons(sv_data, kq)
    return kq


def _parse_api_keys(api_key_str: str) -> list[str]:
    """Tách chuỗi API Key (hỗ trợ nhập nhiều API Key phân cách bởi dấu phẩy, chấm phẩy hoặc khoảng trắng để xoay vòng hạn mức Free Tier)."""
    raw = (api_key_str or "").strip()
    if not raw:
        return []
    keys = [k.strip() for k in re.split(r"[,;\s]+", raw) if k.strip()]
    return keys if keys else [raw]


def _emit_thought_summary(model_name: str, summary_text: str) -> None:
    """Phát tín hiệu suy luận nội bộ (thought summary) của model lên giao diện Tiến trình Suy luận Đa Mô hình."""
    cb = _ACTIVE_THOUGHT_CALLBACK
    if cb is not None and summary_text and summary_text.strip():
        try:
            cb(model_name, summary_text.strip())
        except Exception:
            pass


def _call_gemini_interactions(
    api_key: str,
    input_data,
    schema_dict: dict,
    preferred_model: str = DEFAULT_MODEL,
    timeout_sec: int = 120,
    thinking_level: str = DEFAULT_THINKING_LEVEL,
) -> str:
    """Gọi Gemini Interactions API với thinking_level=high, thinking_summaries=auto, timeout 120s và tự động xoay vòng Key/Model nếu gặp 429/503."""
    api_keys = _parse_api_keys(api_key)
    if not api_keys:
        raise RuntimeError("Thiếu Gemini API Key.")
    models_to_try = [preferred_model] + [m for m in AVAILABLE_MODELS if m != preferred_model]
    url = "https://generativelanguage.googleapis.com/v1beta/interactions"
    last_err = None

    for key_idx, active_key in enumerate(api_keys):
        for model_name in models_to_try:
            exhaust_token = f"{key_idx}:{model_name}" if len(api_keys) > 1 else model_name
            if exhaust_token in _DAILY_EXHAUSTED_MODELS or (len(api_keys) == 1 and model_name in _DAILY_EXHAUSTED_MODELS):
                continue
            cur_thinking = thinking_level or DEFAULT_THINKING_LEVEL
            include_thought_summaries = True
            for attempt in range(2):
                gen_cfg = {"thinking_level": cur_thinking}
                if include_thought_summaries:
                    gen_cfg["thinking_summaries"] = "auto"
                payload = {
                    "model": model_name,
                    "input": input_data,
                    "generation_config": gen_cfg,
                    "response_format": {
                        "type": "text",
                        "mime_type": "application/json",
                        "schema": schema_dict,
                    },
                }
                req = urllib.request.Request(
                    url,
                    data=json.dumps(payload).encode("utf-8"),
                    headers={
                        "Content-Type": "application/json",
                        "x-goog-api-key": active_key,
                        "Api-Revision": "2026-05-20",
                    },
                )
                try:
                    with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
                        body = json.loads(resp.read().decode("utf-8"))
                        steps = body.get("steps", [])
                        thought_chunks = []
                        for step in steps:
                            if step.get("type") == "thought":
                                s_txt = str(step.get("summary") or "").strip()
                                if not s_txt:
                                    s_txt = "".join(c.get("text", "") for c in step.get("content", [])).strip()
                                if s_txt:
                                    thought_chunks.append(s_txt)
                        if thought_chunks:
                            _emit_thought_summary(model_name, " ".join(thought_chunks))
                        out_text = "".join(
                            c.get("text", "")
                            for step in steps
                            if step.get("type") == "model_output"
                            for c in step.get("content", [])
                        )
                        if out_text.strip():
                            return out_text
                        raise RuntimeError(f"Model {model_name} trả về rỗng.")
                except urllib.error.HTTPError as e:
                    err_body = e.read().decode("utf-8", errors="ignore")
                    last_err = f"HTTP {e.code} ({model_name}): {err_body[:220]}"
                    # Nếu model không hỗ trợ thinking_level=high hoặc thinking_summaries=auto thì hạ xuống cấu hình tương thích và thử lại ngay
                    if e.code == 400 and ("thinking" in err_body.lower() or "summaries" in err_body.lower()):
                        if cur_thinking != "low" or include_thought_summaries:
                            cur_thinking = "low"
                            include_thought_summaries = False
                            continue
                    # Nếu vượt quá giới hạn ngày (per day / rpd) thì đánh dấu hết quota ngày và xoay sang model/key dự phòng kế tiếp
                    if e.code == 429 and ("per day" in err_body.lower() or "rpd" in err_body.lower() or "quota" in err_body.lower()):
                        _DAILY_EXHAUSTED_MODELS.add(exhaust_token)
                        if len(api_keys) == 1:
                            _DAILY_EXHAUSTED_MODELS.add(model_name)
                        for other_m in AVAILABLE_MODELS:
                            if other_m.lower() in err_body.lower():
                                _DAILY_EXHAUSTED_MODELS.add(f"{key_idx}:{other_m}" if len(api_keys) > 1 else other_m)
                                if other_m == "gemini-3.6-flash":
                                    _DAILY_EXHAUSTED_MODELS.add(f"{key_idx}:gemini-3.5-flash" if len(api_keys) > 1 else "gemini-3.5-flash")
                                elif other_m == "gemini-3.8-flash":
                                    _DAILY_EXHAUSTED_MODELS.add(f"{key_idx}:gemini-3.7-flash" if len(api_keys) > 1 else "gemini-3.7-flash")
                        break
                    if e.code in (429, 500, 502, 503, 504):
                        time.sleep(3 * (attempt + 1))
                        continue
                    raise RuntimeError(last_err)
                except Exception as e:
                    last_err = f"Lỗi kết nối ({model_name}): {e}"
                    time.sleep(2 * (attempt + 1))
                    continue

    raise RuntimeError(f"Tất cả các model đều quá tải hoặc hết hạn mức API: {last_err}")


def _format_essay_with_line_numbers(part_dict: dict) -> str:
    """Hiển thị bài làm kèm đánh số [Dòng 1], [Dòng 2]... để AI nhìn thấy 100% vị trí sinh viên nhấn Enter xuống hàng."""
    lines = part_dict.get("essay_lines")
    if not lines:
        raw = (part_dict.get("essay") or "").strip()
        lines = [l.strip() for l in raw.splitlines() if l.strip()] if raw else []
    if not lines:
        return "[BỎ TRỐNG]"
    return "\n".join(f"[Dòng {i}] {line}" for i, line in enumerate(lines, 1))


def cham_1_sinh_vien(
    api_key: str,
    sv_data: dict,
    rubric_text: str,
    model_name: str = DEFAULT_MODEL,
) -> tuple[str, KetQuaCham]:
    mssv = str(sv_data.get("mssv", "")).strip()
    p1 = sv_data.get("p1") or {}
    p2 = sv_data.get("p2") or {}
    e1 = (p1.get("essay") or "").strip()
    e2 = (p2.get("essay") or "").strip()
    w1, w2 = count_words(e1), count_words(e2)

    if w1 == 0 and w2 == 0:
        kq_empty = KetQuaCham(
            p1_content=0, p1_org=0, p1_grammar=0, p1_vocab=0,
            p2_content=0, p2_org=0, p2_grammar=0, p2_vocab=0,
            nhan_xet="Sinh viên bỏ trắng cả 2 bài (0 từ).",
        )
        return mssv, ensure_cogv_reasons(sv_data, kq_empty)

    p1_images = p1.get("images") or []
    p2_images = p2.get("images") or []
    pr1 = clean_prompt_text(p1.get("prompt") or "")
    pr2 = clean_prompt_text(p2.get("prompt") or "")
    fn1 = (p1.get("format_note") or "").strip()
    fn2 = (p2.get("format_note") or "").strip()
    e1_lined = _format_essay_with_line_numbers(p1)
    e2_lined = _format_essay_with_line_numbers(p2)
    l1 = p1.get("line_count") or (len(e1.splitlines()) if e1 else 0)
    l2 = p2.get("line_count") or (len(e2.splitlines()) if e2 else 0)

    prompt = f"""Bạn là giảng viên chấm thi môn Anh Văn Căn Bản (Writing - English for Beginners A1) tại trường Đại học Công nghiệp TP.HCM (IUH).
Hãy chấm bài nghiêm túc, chặt tay, khách quan theo đúng Rubric chuẩn của Khoa dưới đây:
{rubric_text}

=== WRITING PART 1 (Mã đề: {p1.get('code', 'N/A')}) ===
- Đề bài gốc (Prompt):
{pr1 or '[Xem hình ảnh đề bài Part 1 đính kèm]'}
- Số ảnh đề bài Part 1 đính kèm: {len(p1_images)}
- Ghi chú trình bày P1 ({l1} dòng): {fn1 or 'Bình thường (đạt yêu cầu trình bày đoạn văn)'}
- Bài làm sinh viên ({w1} từ, {l1} dòng - mỗi [Dòng k] là 1 lần sinh viên xuống hàng):
{e1_lined}

=== WRITING PART 2 (Mã đề: {p2.get('code', 'N/A')}) ===
- Đề bài gốc (Prompt):
{pr2 or '[Xem hình ảnh đề bài Part 2 đính kèm]'}
- Số ảnh đề bài Part 2 đính kèm: {len(p2_images)}
- Ghi chú trình bày P2 ({l2} dòng): {fn2 or 'Bình thường (đạt yêu cầu trình bày đoạn văn)'}
- Bài làm sinh viên ({w2} từ, {l2} dòng - mỗi [Dòng k] là 1 lần sinh viên xuống hàng):
{e2_lined}

QUY TẮC BẮT BUỘC (v7 - GOLD STANDARD 2607AVCB005 & 2607AVCB006):
1. ĐẶT MÌNH VÀO VỊ TRÍ NGƯỜI NHẬN THƯ Ở MỌI TÌNH HUỐNG VIẾT MAIL (CẢ PART 1 & PART 2):
   - Khi đọc bất kỳ bức thư nào (P1 hay P2), hãy tự hỏi dưới góc nhìn người nhận thư: "Đọc xong mail này mình đã hiểu rõ câu trả lời/thông tin chưa, hay vẫn phải hỏi ngược lại người viết (ví dụ: đề hỏi mấy giờ đi mà chỉ nói 7h concert bắt đầu -> 'ủa vậy mấy giờ tụi mình đi?')?". Nếu người nhận vẫn phải thắc mắc hoặc câu văn lệch ý câu hỏi -> trừ điểm Content.
   - Kiểm tra phần kết thư & ký tên ở cuối bài: Nếu thiếu câu kết hoặc thiếu ký tên người viết (kể cả chỉ ghi 'Your friend.' / 'Dear,' mà không có tên) -> trừ 1 điểm Content (tối đa Content = 4).
   - Lưu ý ở Part 1: Nếu sinh viên đặt đủ 2 câu hỏi đúng chủ đề cho bạn nhưng quên dấu `?` ở cuối câu (dùng dấu `.`, ví dụ `What time do you wake up. And what your usually does in the morning.`) thì KHÔNG trừ điểm Content (vẫn cho 5 Content nếu đủ ý). Ngược lại, các câu mời xã giao như `Would you like to my house?` hoặc `OK for you?` KHÔNG được tính là câu hỏi chủ đề của đề bài.
   - Kiểm tra bài dài: KHÔNG tự động trừ điểm chỉ vì bài dài ~90-100 từ nếu sinh viên viết đúng trọng tâm và mạch lạc. CHỈ trừ 1 điểm Content khi bài dài (~80 từ trở lên) VÀ có câu dẫn nhập sáo rỗng (`Today I write an email to talk about...`, `Long time no meet you...`, `The first one...`) hoặc kể lể lạc sang buổi khác.
2. NHẬN XÉT CHI TIẾT DÀNH CHO GIÁM KHẢO KIỂM CHỨNG (`p1_c_reason`, `p1_o_reason`, `p1_g_reason`, `p1_v_reason` và `p2_c_reason`, `p2_o_reason`, `p2_g_reason`, `p2_v_reason`):
   - Người đọc phần nhận xét này là GIÁM KHẢO (Examiner). Tuyệt đối KHÔNG viết nhận xét chung chung kiểu "Còn một số lỗi chia động từ..." hay điểm 3 mà lại ghi "Đạt: Trả lời đầy đủ...".
   - Bắt buộc chỉ rõ: Đạt mấy ý / thiếu ý nào của đề, và TRÍCH DẪN NGUYÊN VĂN cụm từ sai trong dấu ngoặc kép `"..."` (ví dụ: `sai cấu trúc ở "My Mom is cooked it", "It very delicous"`).
   - Nếu bài có từ 2 lỗi cấu trúc ngữ pháp cơ bản trở lên -> chấm ngay Grammar = 3 (hoặc 2 nếu lỗi nặng hệ thống) ngay từ đầu, không chấm 4.
3. Về Organization & Trình bày (Format):
   - Châm chước bỏ qua các lỗi nhỏ về khoảng cách dấu câu; trong viết thư thân mật (informal email), việc chấm câu rồi viết `. Because...` hoặc `. So...` vẫn được chấp nhận (KHÔNG trừ điểm). CHỈ trừ điểm trình bày xuống dòng ở cột Organization nếu có ghi chú "Sinh viên xuống dòng rời rạc sau mỗi câu đơn lẻ".
   - TUYỆT ĐỐI KHÔNG chấm 5 điểm Organization nếu sinh viên mắc lỗi Run-on sentence / Comma splice (dùng dấu phẩy `,` thay cho dấu chấm câu `.` để nối các câu đơn độc lập mà không có liên từ, phải hạ xuống 3).
4. Về Vocabulary:
   - Nếu sinh viên chỉ dùng từ vựng cơ bản, đơn giản (như father, mother, friend, happy, good, nice, fun, big, small, house, bedroom, watch TV, football, badminton, T-shirt, shoes, hat, fruit, bread, milk, water, coffee, Pho, Banh mi...) mà không có cụm từ mô tả đặc sắc: TỐI ĐA CHỈ CHẤM 3 ĐIỂM VOCABULARY (không được cho 4!).
5. Quy tắc Khóa/Phạt điểm:
   - Nếu lạc đề hoàn toàn / học thuộc văn mẫu sai đề (Content = 0, bài >= 12 từ), bao gồm cả lỗi ở Part 2 tự đóng vai người nhận của Part 1 để trả lời cho bài Part 1 của chính mình, hoặc Part 1 chép văn mẫu đặt lịch hẹn `Would you like to... I'm free at...`: chấm 0 - 1 - 1 - 1.
   - Nếu chép nguyên văn đề bài hoặc viết dở dang 1 câu cụt < 12 từ: chấm 0 - 0 - 0 - 0.
   - Nếu không hiểu đề (Content = 1) hoặc nửa đúng đề nửa học tủ (Content = 2): các cột Organization, Grammar, Vocabulary bị phạt trừ 1 điểm so với thực tế (sàn tối thiểu là 1).
"""

    input_contents = []
    for img in p1_images:
        if isinstance(img, dict) and img.get("data") and img.get("mime_type"):
            input_contents.append({"type": "image", "mime_type": img["mime_type"], "data": img["data"]})
    for img in p2_images:
        if isinstance(img, dict) and img.get("data") and img.get("mime_type"):
            input_contents.append({"type": "image", "mime_type": img["mime_type"], "data": img["data"]})
    input_contents.append({"type": "text", "text": prompt})

    out_text = _call_gemini_interactions(
        api_key=api_key,
        input_data=input_contents if len(input_contents) > 1 else prompt,
        schema_dict=KetQuaCham.model_json_schema(),
        preferred_model=model_name,
    )
    kq = KetQuaCham.model_validate_json(out_text)
    kq = _calibrate_scores(sv_data, kq)
    return mssv, kq


def cham_1_lo_sinh_vien(
    api_key: str,
    batch_items: list[tuple[int, str, str, dict]],
    rubric_text: str,
    model_name: str = DEFAULT_MODEL,
) -> dict[str, KetQuaCham]:
    """Chấm gộp 1 lô (tối đa 7 sinh viên) trong 1 lần gọi API để tiết kiệm số lượng request (tránh lỗi 429)."""
    results: dict[str, KetQuaCham] = {}
    need_api_items: list[tuple[int, str, str, dict, int, int]] = []

    for row_idx, mssv, ten, sv_data in batch_items:
        p1 = sv_data.get("p1") or {}
        p2 = sv_data.get("p2") or {}
        e1 = (p1.get("essay") or "").strip()
        e2 = (p2.get("essay") or "").strip()
        w1, w2 = count_words(e1), count_words(e2)

        if w1 == 0 and w2 == 0:
            kq_empty = KetQuaCham(
                p1_content=0, p1_org=0, p1_grammar=0, p1_vocab=0,
                p2_content=0, p2_org=0, p2_grammar=0, p2_vocab=0,
                nhan_xet="Sinh viên bỏ trắng cả 2 bài (0 từ).",
            )
            results[mssv] = ensure_cogv_reasons(sv_data, kq_empty)
        elif (p1.get("images") or []) or (p2.get("images") or []):
            # Nếu câu hỏi có đính kèm ảnh đề bài thì chấm riêng từng SV để gửi kèm ảnh Multimodal
            _, kq_single = cham_1_sinh_vien(api_key, sv_data, rubric_text, model_name=model_name)
            results[mssv] = kq_single
        else:
            need_api_items.append((row_idx, mssv, ten, sv_data, w1, w2))

    if not need_api_items:
        return results

    blocks = []
    for idx, (_, mssv, _, sv_data, w1, w2) in enumerate(need_api_items, 1):
        p1 = sv_data.get("p1") or {}
        p2 = sv_data.get("p2") or {}
        pr1 = clean_prompt_text(p1.get("prompt") or "")
        pr2 = clean_prompt_text(p2.get("prompt") or "")
        fn1 = (p1.get("format_note") or "").strip()
        fn2 = (p2.get("format_note") or "").strip()
        e1 = (p1.get("essay") or "").strip()
        e2 = (p2.get("essay") or "").strip()
        e1_lined = _format_essay_with_line_numbers(p1)
        e2_lined = _format_essay_with_line_numbers(p2)
        l1 = p1.get("line_count") or (len(e1.splitlines()) if e1 else 0)
        l2 = p2.get("line_count") or (len(e2.splitlines()) if e2 else 0)
        blocks.append(
            f"--- SINH VIÊN #{idx} | MSSV: {mssv} ---\n"
            f"[PART 1 - Mã đề: {p1.get('code', 'N/A')}]\n"
            f"• Đề bài P1:\n{pr1}\n"
            f"• Ghi chú trình bày P1 ({l1} dòng): {fn1 or 'Bình thường (đạt yêu cầu trình bày đoạn văn)'}\n"
            f"• Bài làm P1 ({w1} từ, {l1} dòng - mỗi [Dòng k] là 1 lần sinh viên xuống hàng):\n{e1_lined}\n"
            f"[PART 2 - Mã đề: {p2.get('code', 'N/A')}]\n"
            f"• Đề bài P2:\n{pr2}\n"
            f"• Ghi chú trình bày P2 ({l2} dòng): {fn2 or 'Bình thường (đạt yêu cầu trình bày đoạn văn)'}\n"
            f"• Bài làm P2 ({w2} từ, {l2} dòng - mỗi [Dòng k] là 1 lần sinh viên xuống hàng):\n{e2_lined}"
        )

    prompt = (
        "Bạn là giảng viên chấm thi môn Anh Văn Căn Bản (Writing - English for Beginners A1) tại trường Đại học Công nghiệp TP.HCM (IUH).\n"
        "Hãy chấm bài độc lập, nghiêm túc, chặt tay theo đúng chuẩn giảng viên IUH (v7 - Gold Standard 2607AVCB005 & 2607AVCB006) cho từng sinh viên dưới đây:\n\n"
        f"{rubric_text}\n\n"
        "QUY TẮC BẮT BUỘC:\n"
        "1. ĐẶT MÌNH VÀO VỊ TRÍ NGƯỜI NHẬN THƯ Ở MỌI TÌNH HUỐNG VIẾT MAIL (CẢ PART 1 & PART 2):\n"
        "   - Khi đọc bất kỳ bức thư nào (P1 hay P2), hãy tự đặt câu hỏi dưới góc nhìn người nhận thư: 'Đọc xong mail này mình đã hiểu rõ câu trả lời/thông tin chưa, hay vẫn phải hỏi ngược lại người viết (ví dụ: đề hỏi mấy giờ đi mà chỉ nói 7h concert bắt đầu -> ủa vậy mấy giờ tụi mình đi?; hoặc đề hỏi tên phim là gì `Can you let me know the movie title?` mà viết câu vô nghĩa `Can you watch the movie title happy.` -> người nhận hoàn toàn không hiểu tên phim là gì)?'. Nếu người nhận vẫn phải thắc mắc hoặc câu văn vô nghĩa/lệch ý câu hỏi -> bắt buộc trừ điểm Content!\n"
        "   - CẢNH GIÁC THỦ THUẬT CHÉP GẠCH ĐẦU DÒNG ĐỀ BÀI Ở PART 1 HOẶC TỰ TRẢ LỜI CHÉO P1-P2: Nếu đề P1 yêu cầu kể về bản thân mà sinh viên biến toàn bộ thành câu hỏi cho bạn (`What time do you wake up? What do you do in the morning? What do you have for breakfast?`) hoặc ở P2 sinh viên phớt lờ người gửi trên đề P2 để tự đóng vai người nhận của P1 viết thư hồi đáp cho chính bài P1 của mình -> Khóa cứng 0 - 1 - 1 - 1!\n"
        "   - Lưu ý ở Part 1: Nếu sinh viên đặt đủ 2 câu hỏi đúng chủ đề cho bạn nhưng quên dấu `?` ở cuối câu (dùng dấu `.`, ví dụ `What time do you wake up. And what your usually does in the morning.`) thì KHÔNG trừ điểm Content. Ngược lại, các câu mời xã giao như `Would you like to my house?` hoặc `OK for you?` KHÔNG tính là câu hỏi chủ đề.\n"
        "   - Kiểm tra phần kết thư & ký tên ở cuối bài: Nếu thiếu câu kết hoặc thiếu ký tên người viết (kể cả chỉ ghi `Your friend.` / `Dear,` không kèm tên) -> trừ 1 điểm Content (tối đa Content = 4).\n"
        "   - Kiểm tra bài dài: CHỈ trừ 1 điểm Content khi bài dài (~80 từ trở lên) VÀ có câu dẫn sáo rỗng (`Today I write an email to talk about...`, `Long time no meet you...`) hoặc kể lể lạc sang buổi khác. Nếu viết ~90-100 từ nhưng đúng trọng tâm và mạch lạc thì KHÔNG trừ điểm.\n"
        "2. NHẬN XÉT CHI TIẾT DÀNH CHO GIÁM KHẢO KIỂM CHỨNG (`p1_c_reason`, `p1_o_reason`, `p1_g_reason`, `p1_v_reason` và `p2_c_reason`, `p2_o_reason`, `p2_g_reason`, `p2_v_reason`):\n"
        "   - Người đọc là GIÁM KHẢO (Examiner). Không viết nhận xét chung chung hoặc mâu thuẫn với điểm số (tuyệt đối không để điểm 3 mà ghi 'Đạt: Trả lời đầy đủ...').\n"
        "   - Phải chỉ rõ thiếu/lệch ý nào của đề và TRÍCH DẪN NGUYÊN VĂN cụm từ sai ngữ pháp/chính tả trong ngoặc kép `\"...\"`.\n"
        "   - Nếu bài có từ 2 lỗi cấu trúc ngữ pháp cơ bản trở lên -> chấm ngay Grammar = 3 (hoặc 2 nếu lỗi nặng hệ thống) ngay từ đầu, không chấm 4.\n"
        "3. Về Organization & Trình bày (Format):\n"
        "   - Châm chước bỏ qua các lỗi nhỏ về khoảng cách dấu câu; trong viết thư thân mật (informal email), việc chấm câu rồi viết `. Because...` hoặc `. So...` vẫn được chấp nhận (KHÔNG trừ điểm). CHỈ trừ điểm trình bày xuống dòng ở cột Organization nếu có ghi chú 'Sinh viên xuống dòng rời rạc sau mỗi câu đơn lẻ'.\n"
        "   - TUYỆT ĐỐI KHÔNG chấm 5 điểm Organization nếu sinh viên mắc lỗi Run-on sentence / Comma splice (dùng dấu phẩy `,` thay cho dấu chấm câu `.` để nối các câu đơn độc lập mà không có liên từ, phải hạ xuống 3).\n"
        "4. Về Vocabulary:\n"
        "   - Nếu sinh viên chỉ dùng từ vựng cơ bản, đơn giản (như father, mother, friend, happy, good, nice, fun, big, small, house, bedroom, watch TV, football, badminton, T-shirt, shoes, hat, fruit, bread, milk, water, coffee, Pho, Banh mi...) mà không có cụm từ mô tả đặc sắc: TỐI ĐA CHỈ CHẤM 3 ĐIỂM VOCABULARY (không được cho 4!).\n"
        "5. Quy tắc Khóa/Phạt điểm:\n"
        "   - Nếu lạc đề hoàn toàn / học thuộc văn mẫu sai đề (Content = 0, bài >= 12 từ): chấm 0 - 1 - 1 - 1.\n"
        "   - Nếu chép nguyên văn đề bài hoặc viết dở dang 1 câu cụt < 12 từ: chấm 0 - 0 - 0 - 0.\n"
        "   - Nếu không hiểu đề (Content = 1) hoặc nửa đúng đề nửa học tủ (Content = 2): các cột Organization, Grammar, Vocabulary bị phạt trừ 1 điểm so với thực tế (sàn tối thiểu là 1).\n\n"
        + "\n\n".join(blocks)
    )

    out_text = _call_gemini_interactions(
        api_key=api_key,
        input_data=prompt,
        schema_dict=KetQuaChamLo.model_json_schema(),
        preferred_model=model_name,
        timeout_sec=60,
    )
    lo_kq = KetQuaChamLo.model_validate_json(out_text)
    map_lo = {normalize_mssv(item.mssv): item for item in lo_kq.danh_sach}

    for _, mssv, _, sv_data, w1, w2 in need_api_items:
        item = map_lo.get(mssv)
        if item is not None:
            kq = KetQuaCham(
                nhan_xet=item.nhan_xet,
                p1_content=0 if w1 < 5 else item.p1_content,
                p1_org=0 if w1 < 5 else item.p1_org,
                p1_grammar=0 if w1 < 5 else item.p1_grammar,
                p1_vocab=0 if w1 < 5 else item.p1_vocab,
                p1_c_reason=item.p1_c_reason,
                p1_o_reason=item.p1_o_reason,
                p1_g_reason=item.p1_g_reason,
                p1_v_reason=item.p1_v_reason,
                p2_content=0 if w2 < 5 else item.p2_content,
                p2_org=0 if w2 < 5 else item.p2_org,
                p2_grammar=0 if w2 < 5 else item.p2_grammar,
                p2_vocab=0 if w2 < 5 else item.p2_vocab,
                p2_c_reason=item.p2_c_reason,
                p2_o_reason=item.p2_o_reason,
                p2_g_reason=item.p2_g_reason,
                p2_v_reason=item.p2_v_reason,
            )
            kq = _calibrate_scores(sv_data, kq)
            results[mssv] = kq
        else:
            # Dự phòng nếu thiếu 1 SV trong lô thì chấm lẻ SV đó
            _, kq_fallback = cham_1_sinh_vien(api_key, sv_data, rubric_text, model_name=model_name)
            results[mssv] = kq_fallback

    return results


_FALLBACK_GOLD_EXAMPLES_BY_CODE: dict[str, list[dict]] = {
    "AVCB.FT.W.P1.L001": [{"ten": "Phạm Nguyễn Bảo Nhi", "essay": "Hi Mary, / To relax in my city, I often go to Crescent Mall. I usually walk around and buy some fresh fruit and snacks. I like this place because it is very big and beautiful. Where do you usually go to relax? What do you do there? / Best, / Nhi", "c": 5, "o": 5, "g": 5, "v": 4, "rc": "Đạt trọn vẹn (5đ): Đáp ứng đủ 3 ý thông tin & 2 câu hỏi cho bạn, đúng vai người viết thư (49 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Đạt: Viết thành đoạn văn mạch lạc, dùng từ nối và ngắt câu chuẩn.", "rg": "Đạt: Kiểm soát tốt cấu trúc câu, không có lỗi ngữ pháp.", "rv": "Đạt: Từ vựng tự nhiên, đúng chủ đề, không sai chính tả."}],
    "AVCB.FT.W.P1.L002": [{"ten": "Nguyễn Hoàng Lan Anh", "essay": "Hi, Chau / Every day, I wake up at 6:00a.m . In the morning, I have breakfast at 7:00a.m . I go to school at 8:00a.m . In the afternoon, I have for lunch with my friends at 13:00p.m . I usually have for lunch with my friends because it is at school. We are ate rice, eggs, beef, in canteen. I like to have for lunch because it is relax and talk about with my friends.", "c": 2, "o": 2, "g": 2, "v": 2, "rc": "Lỗi: Quên đặt câu hỏi cho bạn và viết lan man nội dung không yêu cầu.", "ro": "Bị phạt trừ điểm do lỗi Content và lặp cấu trúc đầu câu liên tục.", "rg": "Nhiều lỗi ngữ pháp nặng (2đ): Sai cấu trúc động từ và trật tự từ ở nhiều câu trong bài.", "rv": "Bị phạt trừ điểm do lỗi Content."}],
    "AVCB.FT.W.P1.L003": [{"ten": "Phạm Thị Bình Chi", "essay": "Hi, Lan / How are you? / I want to talk about my activity in the morning. / In the morning, I usually get up at 6:00. / Sometimes, I eat Bun Bo Hue for breakfast. Then I go to school at 7:00. / I study from 7:00 to 11:00. / And you. What time do you wake up? And what do you do in the morning?", "c": 4, "o": 3, "g": 5, "v": 3, "rc": "Đạt đủ ý chính nhưng thiếu kết thư/ký tên (4đ): Đã trả lời các yêu cầu của đề (65 từ) nhưng cuối thư kết thúc đột ngột (thiếu dòng kết thư hoặc thiếu ký tên người viết, -1đ).", "ro": "Lỗi trình bày đoạn văn (3đ): Xuống dòng rời rạc sau mỗi câu đơn lẻ thay vì viết thành đoạn văn email liền mạch.", "rg": "Ngữ pháp tốt (5đ): Kiểm soát tốt cấu trúc câu đơn và câu hỏi, chia động từ và dùng giới từ chuẩn xác.", "rv": "Từ vựng cơ bản (3đ): Chủ yếu dùng các từ đơn giản quen thuộc hoặc lấy lại từ khóa từ đề bài, chưa có cụm từ mở rộng."}],
    "AVCB.FT.W.P1.L004": [{"ten": "Trần Phương Uyên", "essay": "Hi Hai, / I'm very miss you. Today, is fine for me. After school, I usually go to the cinema with my friend. Then, I go home at 6:00. In the evening, I sometimes do my homework and cooking dinner for family. I always go to bed at 11:00 because I feel tired for a long day. Let me know what you do go after school, and what / time you go to bed. / Reply me soon. / See you soon! / Phuong Uyen", "c": 4, "o": 5, "g": 4, "v": 4, "rc": "Đạt phần lớn yêu cầu (4đ): Trả lời tốt các ý chính, có mở/kết thư nhưng còn 1 chi tiết diễn đạt chưa rõ cho người nhận (-1đ).", "ro": "Bố cục tốt (5đ): Tổ chức đoạn văn email mạch lạc, đầy đủ mở - thân - kết thư, dùng từ nối và ngắt câu chuẩn xác.", "rg": "Ngữ pháp khá (4đ): Cấu trúc câu chính vững, chỉ có lỗi nhẹ ở \"what you do go\".", "rv": "Từ vựng khá (4đ): Dùng từ vựng đúng chủ đề, diễn đạt rõ nghĩa và không sai chính tả."}],
    "AVCB.FT.W.P1.L005": [{"ten": "Nguyễn Tuấn Phong", "essay": "Hi, An. / My name is Phong. / How are you? / At weekend, I like learn and do my homework because it very happy and comfortable. I usually go to the walk on the park with my family. The park is near market or I usually go to the restaurant with my friends because it help relax and comfortable. / What your friend does on the weekend? / Best / Phong", "c": 4, "o": 3, "g": 3, "v": 3, "rc": "Đạt đủ ý thông tin bản thân, câu hỏi còn lộn ngôi.", "ro": "Mắc lỗi comma splice và xuống dòng rời rạc.", "rg": "Ngữ pháp trung bình (3đ): Đạt câu đơn cơ bản nhưng sai cấu trúc ở \"it very happy\".", "rv": "Từ vựng cơ bản."}],
    "AVCB.FT.W.P1.L006": [{"ten": "Nguyễn Phước Hiếu", "essay": "Hi Dat, / I write this email to talk about my favourite food. My favourite food is Bun Bo. It is a beef noodles sup. I like Bun Bo because its very delicious. I eat it three or four times a week. I usually eat it in Huong Restaurant near my university. I has some question for you. What are your favourite food? And where you like to eat it? / Your friend, / Hieu", "c": 5, "o": 4, "g": 3, "v": 4, "rc": "Đáp ứng 100% yêu cầu về thông tin món bún bò và 2 câu hỏi cho bạn, chào/kết thư đầy đủ.", "ro": "Mạch văn mạch lạc, xếp câu hợp lý.", "rg": "Ngữ pháp trung bình (3đ): Đạt câu đơn cơ bản nhưng sai cấu trúc ở \"I has\", \"where you like\".", "rv": "Từ vựng khá tốt (beef noodles sup, delicious, university) bù trừ cho lỗi chính tả nhỏ."}],
    "AVCB.FT.W.P1.L007": [{"ten": "Nguyễn Thị Tú Quỳnh", "essay": "Hi Luong, / What your favourite sport? I'm so curious. I like play tennis because it is have fun, help me stay healthy and improve skill. How often you play? I'm often play it at weekend with my friends, sometimes I play with my family. How about you? Who you play it with? I want to play with you. Where you can play that sport? I usually play it at The Earth. / Reply me soon. / See you. / Quynh", "c": 5, "o": 4, "g": 3, "v": 4, "rc": "Đạt trọn vẹn (5đ): Đáp ứng đủ 3 ý thông tin & 2 câu hỏi cho bạn, đúng vai người viết thư (77 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục khá (4đ): Trình tự các ý hợp lý, đúng chuẩn email nhưng lặp cấu trúc mở đầu câu bằng 'I / I am / I'm' (5 câu).", "rg": "Lỗi ngữ pháp (3/5): Sai cấu trúc \"I'm often\", \"What your favourite\", \"like play\".", "rv": "Từ vựng khá (4đ): Từ vựng phù hợp ngữ cảnh giao tiếp, có cụm từ tốt (\"curious\", \"stay healthy\")."}],
    "AVCB.FT.W.P1.L008": [{"ten": "Phạm Minh Long", "essay": "Hi, Vinh! / I will tell you about my house. It's apartment with 5 room. It has 2 bed room,1 living room,1 bath room and 1 kitchen. My favorite room it living room. Because it has a TV and my family usually watch TV every evening. / Where do you live Vinh? / What is your favorite room? / It's your friend, Minh Long", "c": 4, "o": 4, "g": 3, "v": 3, "rc": "Đạt phần lớn yêu cầu (4đ): Trả lời tốt các ý chính, có mở/kết thư nhưng còn 1 chi tiết diễn đạt chưa rõ cho người nhận (-1đ).", "ro": "Bố cục khá (4đ): Đầy đủ mở - thân - kết thư, trình tự trả lời hợp lý.", "rg": "Ngữ pháp trung bình (3đ): Còn một số lỗi chia động từ, thiếu to be ở \"My favorite room it living room\".", "rv": "Từ vựng cơ bản (3đ): Chủ yếu dùng các từ đơn giản quen thuộc."}],
    "AVCB.FT.W.P1.L009": [{"ten": "Nguyễn Ngọc Quỳnh Như", "essay": "Hi Tom / How are you ? I hope you are well. Today I write for you to give information about. The first one my family there are 4 people and there are my father , my mother, me and my sister. / The last question I like have lunch with my family after school and go to the park at weekend . Because it make me feel happy and relax . And you can tell me how many people there are in your family . / Who they are , and when your have free time than your like to do with your family . Do you like watch TV with my family ? / I hope you reply . See you soon . / Nhu", "c": 4, "o": 4, "g": 2, "v": 3, "rc": "Đạt ý nhưng viết dài lan man (112 từ, có câu dẫn sáo rỗng, -1đ).", "ro": "Bố cục khá (4đ): Đầy đủ mở - thân - kết thư.", "rg": "Nhiều lỗi ngữ pháp nặng (2đ): Sai cấu trúc câu ở \"it make\", \"like watch\", làm câu văn tối nghĩa.", "rv": "Từ vựng cơ bản (3đ)."}],
    "AVCB.FT.W.P1.L010": [{"ten": "Lương Thị Thu Nga", "essay": "Hi Ngan, / I live in Ho Chi Minh city. It's a big and noisy. Ho Chi Minh city has many places. For example: Ben Thanh market, zoo, park, and cafe, but I love going to the cinema on the weekend because it help me reduce stress and relax my mind. Moreover, people very kind and friendly. And you ? Where do you live city or hometown? Can you talk me a city special ? / Reply me soon / See you / Nga", "c": 5, "o": 4, "g": 4, "v": 5, "rc": "Đạt trọn vẹn (5đ): Đáp ứng đủ 3 ý thông tin & 2 câu hỏi cho bạn, đúng vai người viết thư (78 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục khá (4đ): Đầy đủ mở - thân - kết thư.", "rg": "Ngữ pháp khá (4đ): Các câu đơn và câu hỏi đúng cấu trúc chính, chỉ có 1–2 lỗi nhỏ.", "rv": "Đạt cụm từ tốt (reduce stress, relax my mind)."}],
    "AVCB.FT.W.P2.L001": [{"ten": "Nguyễn Thị Bích Nguyên", "essay": "Hi Anna, / Thank you for your invitation, / However, I am sorry. I am going to study English for my exam on monday. So, I am free on sunday. / Would you like to play mini game when we eat seafood in your garden? / Let's meet together in your garden / Seen you soon, / Thanh.", "c": 1, "o": 2, "g": 3, "v": 3, "rc": "Không đạt mục đích giao tiếp (1đ): Người nhận không hiểu được thông điệp trả lời, sai ngữ cảnh xưng hô.", "ro": "Bố cục yếu (2đ): Các câu lắp ghép rời rạc, thiếu mạch liên kết tự nhiên giữa các ý trong thư.", "rg": "Ngữ pháp trung bình (3đ).", "rv": "Từ vựng cơ bản & có lỗi chính tả (3đ): Sai chính tả ở \"monday\", \"sunday\"."}],
    "AVCB.FT.W.P2.L002": [{"ten": "Đào Minh Triết", "essay": "Hi Tom / I also love to come. Let's meet at 10 o'clock in the evening, at Quang Truong Lam Vien Galaxy cinema on Nguyen Van Bao Strees. Movie title is Nguoi Nhen Quay Tro Lai. I also have dinner at the restaurant near the cinema after the movie. It's good ideal. / See you soon / Triet", "c": 5, "o": 3, "g": 4, "v": 4, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 3 câu hỏi của Tom, đúng vai người viết thư (55 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục trung bình (3đ): Có khung thư cơ bản nhưng cách sắp xếp và chuyển ý giữa các câu trả lời còn rời rạc.", "rg": "Ngữ pháp khá (4đ): Các câu đơn đúng cấu trúc chính.", "rv": "Từ vựng khá (4đ)."}],
    "AVCB.FT.W.P2.L003": [{"ten": "Nguyễn Ngọc Quỳnh Như", "essay": "Hi Mark / I am happy because you love it .We will leave on Saturday at 9 o'clock on Ho Dau Tien lake I have free time. Is it that Ok . I will bring some fruit with you and bread . I need hat and water for the trip / I will feel happy and relax the trip. / Reply soon / See you. / Nhu", "c": 5, "o": 4, "g": 3, "v": 3, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 3 câu hỏi của Mark, đúng vai người viết thư (60 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục khá (4đ).", "rg": "Lỗi ngữ pháp (3/5): Sai cấu trúc \"it that Ok\".", "rv": "Từ vựng cơ bản (3đ)."}],
    "AVCB.FT.W.P2.L004": [{"ten": "Trần Phương Uyên", "essay": "Hi Lucy, / I'm excited about the beach trip, too. I'm so happy because you can go with me. Yes, I very like take photos. I'm going to bring a camera for a trip. For the picnic on the beach, I want to eat some fast food and talk together. Then, I want to see the sunset. / See you soon. / Phuong Uyen", "c": 5, "o": 5, "g": 4, "v": 5, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 2 câu hỏi của Lucy, đúng vai người viết thư (61 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục tốt (5đ): Tổ chức đoạn văn email mạch lạc.", "rg": "Ngữ pháp khá (4đ): Cấu trúc câu chính vững, chỉ có lỗi nhẹ ở \"I very like\".", "rv": "Từ vựng tốt (5đ): Dùng từ vựng phong phú, tự nhiên đúng chủ đề."}],
    "AVCB.FT.W.P2.L005": [{"ten": "Phan Thị Lan Anh", "essay": "Hi Emma / Would you like to come with me go to the party ? The party start at 7 p.m. Is that OK for you?. Would you like to eat snack and drink cola when the party start? you can help me prepare the gif for the party. It's on Tan Son Nhat the party, it's on Nguyen Van Bao street, across Industrial Univerity of Ho Chi Minh city / see you soon / Lan Anh", "c": 4, "o": 3, "g": 3, "v": 3, "rc": "Trả lời được giờ tiệc và đồ mang theo, nhưng nhầm vai giao tiếp ở câu đầu (hỏi mời ngược lại) (-1 điểm).", "ro": "Có lỗi Comma splice nhẹ và xếp câu hỏi đan xen chưa thực sự mạch lạc.", "rg": "Ngữ pháp trung bình (3đ): Còn lỗi chia động từ \"The party start\".", "rv": "Từ vựng cơ bản (party, snack, cola, gift, street)."}],
    "AVCB.FT.W.P2.L006": [{"ten": "Nguyễn Thị Thúy Ngân", "essay": "Hi Anna, / We are going to the concert at five o'clock this weekend. The concert on Nguyen Van Nghi Street, Ho Chi Minh city. And I bought the tickets for us. Before the concert, I want to have a drink at the coffee shop near our school. / See you soon, / Ngan.", "c": 5, "o": 5, "g": 4, "v": 3, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 3 câu hỏi của Anna, đúng vai người viết thư (51 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục tốt (5đ): Tổ chức đoạn văn email mạch lạc.", "rg": "Ngữ pháp khá (4đ): Các câu đơn đúng cấu trúc chính, thiếu to be ở 1 câu.", "rv": "Từ vựng cơ bản (3đ)."}],
    "AVCB.FT.W.P2.L007": [{"ten": "Nguyễn Trần Trà Mi", "essay": "Hi Ben, / Thank you for your invitation. / I'm so happy. I like go picnic with you is talking and laughting together. We can meet at Industrial University of Ho Chi Minh City and start go to the trip. My mother doesn't go to take the park because her is busy homework at home. We are going to travel by a car. Oh great! Then, you can bring some food or drinks. Looking forward to relaxing with you! / See you soon. / Mi.", "c": 4, "o": 3, "g": 3, "v": 3, "rc": "Đạt ý nhưng viết dài lan man (81 từ, -1đ).", "ro": "Bố cục trung bình (3đ).", "rg": "Ngữ pháp trung bình (3đ): Đạt câu đơn cơ bản nhưng sai cấu trúc ở \"like go\", \"her is busy\".", "rv": "Lỗi chính tả/dùng từ: \"laughting\"; từ vựng còn lại cơ bản."}],
    "AVCB.FT.W.P2.L008": [{"ten": "Nguyễn Hoàng Nam", "essay": "Hi Sophie / I'm so glad you can join me. Yes, i will buy the tickets online tonight so we don't have to wait in line. Let's meet at 9:00 AM at the main entrance of the museum. You don't need to bring anything, just yourself / See you Saturday, / Nam", "c": 5, "o": 4, "g": 4, "v": 5, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 3 câu hỏi của Sophie, đúng vai người viết thư (50 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục khá (4đ).", "rg": "Ngữ pháp khá (4đ).", "rv": "Từ vựng tốt (5đ): Dùng từ vựng phong phú, tự nhiên đúng chủ đề."}],
    "AVCB.FT.W.P2.L009": [{"ten": "Nguyễn Phi Hùng", "essay": "Hi Rachel / Thank you for your information. I can play football match on Sunday because i am free time and it also can help me relax and reduce tress after a long day of studying . We can meet at 4 p.m. I need you bring water ,money and phone to play game with me . I do not want to go to the club because I am very busy and I do not like the club. / See you soon, / Hung", "c": 4, "o": 4, "g": 3, "v": 4, "rc": "Trả lời đầy đủ các câu hỏi của đề, bài viết hơi dài (83 từ, -1đ).", "ro": "Mạch văn rõ ràng, logic.", "rg": "Ngữ pháp trung bình (3đ): Đạt câu đơn cơ bản nhưng sai cấu trúc ở \"i am free time\".", "rv": "Từ vựng khá tốt 'reduce tress after a long day'."}],
    "AVCB.FT.W.P2.L010": [{"ten": "Nguyễn Thị Tú Quỳnh", "essay": "Hi John, / I'm so happy. I'm going to walk to the bus station at 6 p.m. I'm going to buy pink dress because I need it for my birthday and a black coat for this winter. I want to eat food at Sun World Restaurant because it is delicious. / See you soon! / Quynh", "c": 5, "o": 4, "g": 4, "v": 4, "rc": "Đạt trọn vẹn (5đ): Đáp ứng trả lời trọn vẹn 3 câu hỏi của John, đúng vai người viết thư (54 từ), có lời chào và kết thư/ký tên đầy đủ.", "ro": "Bố cục khá (4đ).", "rg": "Ngữ pháp khá (4đ).", "rv": "Đạt: Từ vựng tự nhiên, đúng chủ đề, không sai chính tả."}],
}
_GOLD_STANDARD_CACHE: dict[str, list[dict]] | None = None


def _load_gold_standard_examples_by_code() -> dict[str, list[dict]]:
    """Nạp bộ bài mẫu chuẩn (Gold Standard Few-Shot) cho đủ 20 mã đề (P1.L001..L010 & P2.L001..L010) từ 2607AVCB005 & 2607AVCB006."""
    global _GOLD_STANDARD_CACHE
    if _GOLD_STANDARD_CACHE is not None:
        return _GOLD_STANDARD_CACHE

    by_code: dict[str, list[dict]] = {
        k: [dict(item) for item in v] for k, v in _FALLBACK_GOLD_EXAMPLES_BY_CODE.items()
    }
    raw_pool: dict[str, list[dict]] = {}

    candidate_roots = [get_app_dir(), os.path.dirname(os.path.dirname(os.path.abspath(__file__)))]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and meipass not in candidate_roots:
        candidate_roots.append(meipass)

    for root_dir in candidate_roots:
        for rel_path in (
            os.path.join("2607AVCB005", "bai_lam_iuh_2607AVCB005_review.json"),
            os.path.join("2607AVCB006", "bai_lam_iuh_2607AVCB006_review.json"),
        ):
            full_p = os.path.join(root_dir, rel_path)
            if not os.path.exists(full_p):
                continue
            try:
                with open(full_p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                items = data.get("items") if isinstance(data, dict) else []
                for it in items or []:
                    if it.get("absent"):
                        continue
                    for pk in ("p1", "p2"):
                        p_obj = it.get(pk) or {}
                        code = str(p_obj.get("code") or "").strip()
                        essay = str(p_obj.get("essay") or "").strip()
                        if not code.startswith("AVCB") or not essay or "[VẮNG" in essay:
                            continue
                        raw_pool.setdefault(code, []).append({
                            "ten": str(it.get("ten") or "").strip(),
                            "essay": " / ".join(ln.strip() for ln in essay.splitlines() if ln.strip()),
                            "c": int(p_obj.get("c", 0)),
                            "o": int(p_obj.get("o", 0)),
                            "g": int(p_obj.get("g", 0)),
                            "v": int(p_obj.get("v", 0)),
                            "rc": str(p_obj.get("rc") or "").strip(),
                            "ro": str(p_obj.get("ro") or "").strip(),
                            "rg": str(p_obj.get("rg") or "").strip(),
                            "rv": str(p_obj.get("rv") or "").strip(),
                        })
            except Exception:
                continue

    for code, samples in raw_pool.items():
        s_desc = sorted(samples, key=lambda x: (x["c"] + x["o"] + x["g"] + x["v"], x["c"]), reverse=True)
        non_zero = [x for x in samples if (x["c"] + x["o"] + x["g"] + x["v"]) > 0] or samples
        s_asc = sorted(non_zero, key=lambda x: (x["c"] + x["o"] + x["g"] + x["v"], x["c"]))
        picked = [s_desc[0]]
        if len(s_desc) > 2:
            mid = s_desc[len(s_desc) // 2]
            if mid["essay"] not in [p["essay"] for p in picked]:
                picked.append(mid)
        if s_asc[0]["essay"] not in [p["essay"] for p in picked]:
            picked.append(s_asc[0])
        by_code[code] = picked[:3]

    _GOLD_STANDARD_CACHE = by_code
    return by_code


def _format_gold_examples_for_code(code: str, is_part1: bool = True, max_examples: int = 3) -> str:
    """Định dạng 1-3 bài mẫu chuẩn giảng viên (Gold Standard) của đúng mã đề `code` để AI bám sát thang điểm."""
    gold_map = _load_gold_standard_examples_by_code()
    examples = gold_map.get((code or "").strip())
    if not examples:
        prefix = "AVCB.FT.W.P1." if is_part1 else "AVCB.FT.W.P2."
        for k, v in gold_map.items():
            if k.startswith(prefix) and v:
                examples = v[:1]
                break
    if not examples:
        return "(Chấm theo đúng Rubric chuẩn giảng viên IUH)."

    lines = []
    for idx, ex in enumerate(examples[: max(1, int(max_examples))], 1):
        lines.append(
            f"  [Mẫu chuẩn #{idx} của mã đề {code}] -> Điểm chuẩn GV: C={ex['c']}, O={ex['o']}, G={ex['g']}, V={ex['v']}\n"
            f"  • Bài viết SV: \"{ex['essay']}\"\n"
            f"  • Lý do chuẩn GV: C: {ex['rc']} | O: {ex['ro']} | G: {ex['rg']} | V: {ex['rv']}"
        )
    return "\n".join(lines)


def _extract_p1_cross_ref_note(sv_data: dict) -> str:
    """Trích xuất dòng chào đầu thư và ký tên cuối thư ở Part 1 của sinh viên để đối chiếu chéo khi chấm Part 2."""
    p1 = sv_data.get("p1") or {}
    e1 = (p1.get("essay") or "").strip()
    if not e1:
        return "Part 1 bỏ trống."
    lines = [ln.strip() for ln in e1.splitlines() if ln.strip()]
    first_ln = lines[0][:60] if lines else ""
    last_ln = lines[-1][:40] if len(lines) > 1 else ""
    p1_code = p1.get("code") or "N/A"
    return f"Ở Part 1 ({p1_code}), SV này mở đầu '{first_ln}' và kết thúc '{last_ln}'."


def _build_prompt_clusters_for_part(
    tasks: list[tuple[int, str, str, dict]],
    part_key: str,
    max_clusters: int = MAX_PROMPT_CLUSTERS_PER_PART,
) -> list[list[tuple[str, list[tuple[int, str, str, dict]]]]]:
    """
    Gom nhóm ngầm trong bộ nhớ các bài làm thuộc `part_key` ('p1' hoặc 'p2') theo từng Mã đề (L001..L010),
    sau đó đóng gói thành tối đa `max_clusters` (mặc định 2) cụm cân bằng nhau.
    Tuyệt đối không làm thay đổi danh sách `tasks` gốc của giao diện.
    """
    by_code: dict[str, list[tuple[int, str, str, dict]]] = {}
    for row_idx, mssv, ten, sv_data in tasks:
        p1_imgs = (sv_data.get("p1") or {}).get("images") or []
        p2_imgs = (sv_data.get("p2") or {}).get("images") or []
        if p1_imgs or p2_imgs:
            continue
        part_obj = sv_data.get(part_key) or {}
        essay = (part_obj.get("essay") or "").strip()
        if count_words(essay) == 0:
            continue
        code = str(part_obj.get("code") or f"{part_key.upper()}_DEFAULT").strip()
        by_code.setdefault(code, []).append((row_idx, mssv, ten, sv_data))

    if not by_code:
        return []

    sorted_groups = sorted(by_code.items(), key=lambda kv: kv[0])
    if len(sorted_groups) <= 1 or max_clusters <= 1:
        return [sorted_groups]

    # Chia danh sách nhóm mã đề thành tối đa max_clusters (2) cụm sao cho số lượng bài làm giữa các cụm cân bằng nhất
    total_essays = sum(len(st_list) for _, st_list in sorted_groups)
    target_per_cluster = total_essays / float(min(max_clusters, len(sorted_groups)))

    clusters: list[list[tuple[str, list[tuple[int, str, str, dict]]]]] = []
    cur_cluster: list[tuple[str, list[tuple[int, str, str, dict]]]] = []
    cur_count = 0

    for idx, group in enumerate(sorted_groups):
        remaining_groups = len(sorted_groups) - idx
        remaining_clusters_allowed = max_clusters - len(clusters) - 1
        cur_cluster.append(group)
        cur_count += len(group[1])
        if (
            remaining_clusters_allowed > 0
            and remaining_groups > remaining_clusters_allowed
            and cur_count >= target_per_cluster * 0.85
        ):
            clusters.append(cur_cluster)
            cur_cluster = []
            cur_count = 0

    if cur_cluster:
        clusters.append(cur_cluster)

    return clusters[:max_clusters]


def cham_1_cum_ma_de(
    api_key: str,
    cluster_groups: list[tuple[str, list[tuple[int, str, str, dict]]]],
    part_key: str,
    rubric_text: str,
    model_name: str = DEFAULT_MODEL,
) -> dict[str, KetQuaCham1PhanSV]:
    """
    Chấm 1 cụm mã đề của Part 1 hoặc Part 2 (tối đa 5 mã đề cùng loại trong 1 request) với thinking_level='high'
    và gắn kèm bài mẫu Gold Standard của đúng mã đề đó ngay dưới mỗi đề bài.
    """
    is_part1 = part_key == "p1"
    part_label = "WRITING PART 1" if is_part1 else "WRITING PART 2"
    section_blocks: list[str] = []
    expected_items: list[tuple[str, dict, int]] = []

    for sec_idx, (code, st_list) in enumerate(cluster_groups, 1):
        sample_part = (st_list[0][3].get(part_key) or {}) if st_list else {}
        pr_text = clean_prompt_text(sample_part.get("prompt") or "")
        gold_text = _format_gold_examples_for_code(code, is_part1=is_part1)

        sv_blocks: list[str] = []
        for st_idx, (_, mssv, _, sv_data) in enumerate(st_list, 1):
            p_obj = sv_data.get(part_key) or {}
            essay_raw = (p_obj.get("essay") or "").strip()
            w_cnt = count_words(essay_raw)
            l_cnt = p_obj.get("line_count") or (len(essay_raw.splitlines()) if essay_raw else 0)
            fn_note = (p_obj.get("format_note") or "").strip() or "Bình thường (đạt yêu cầu trình bày đoạn văn)"
            essay_lined = _format_essay_with_line_numbers(p_obj)
            cross_ref_line = (
                f"• Đối chiếu chéo Part 1 của SV {mssv}: {_extract_p1_cross_ref_note(sv_data)}\n"
                if not is_part1
                else ""
            )
            sv_blocks.append(
                f"--- BÀI LÀM #{st_idx} (Mã đề {code}) | MSSV: {mssv} ---\n"
                f"{cross_ref_line}"
                f"• Ghi chú trình bày ({l_cnt} dòng): {fn_note}\n"
                f"• Bài làm sinh viên ({w_cnt} từ, {l_cnt} dòng - mỗi [Dòng k] là 1 lần sinh viên xuống hàng):\n{essay_lined}"
            )
            expected_items.append((mssv, sv_data, w_cnt))

        section_blocks.append(
            f"====================================================================\n"
            f"=== NGĂN MÃ ĐỀ {part_label} #{sec_idx}: {code} ({len(st_list)} sinh viên) ===\n"
            f"====================================================================\n"
            f"[ĐỀ BÀI GỐC CỦA MÃ ĐỀ {code}]:\n{pr_text}\n\n"
            f"[BÀI MẪU CHUẨN GIẢNG VIÊN (GOLD STANDARD FEW-SHOT) CỦA ĐÚNG MÃ ĐỀ {code}]:\n{gold_text}\n\n"
            f"[DANH SÁCH BÀI LÀM SINH VIÊN THUỘC MÃ ĐỀ {code}]:\n" + "\n\n".join(sv_blocks)
        )

    prompt = (
        f"Bạn là giảng viên chấm thi môn Anh Văn Căn Bản (Writing - English for Beginners A1) tại trường Đại học Công nghiệp TP.HCM (IUH).\n"
        f"Nhiệm vụ: Chấm chuyên sâu **{part_label}** theo từng NGĂN MÃ ĐỀ đã được phân loại sẵn dưới đây (Chuẩn giảng viên IUH v7 - Gold Standard 2607AVCB005 & 2607AVCB006).\n\n"
        f"{rubric_text}\n\n"
        f"QUY TẮC BẮT BUỘC KHI CHẤM CỤM MÃ ĐỀ {part_label}:\n"
        f"1. ĐỐI CHIẾU TRỰC TIẾP VỚI ĐỀ BÀI GỐC & BÀI MẪU GOLD STANDARD CỦA ĐÚNG NGĂN MÃ ĐỀ ĐÓ:\n"
        f"   - Mỗi ngăn mã đề đều có Đề bài gốc và Bài mẫu chuẩn giảng viên (Gold Standard) của chính mã đề đó. Hãy so sánh bài làm của từng sinh viên với Đề bài gốc và Bài mẫu chuẩn cùng mã đề để chấm 4 cột `content`, `org`, `grammar`, `vocab` (0-5) sát 100% tay chấm của giảng viên.\n"
        f"   - Đặt mình vào vị trí người nhận thư: Nếu người nhận đọc xong vẫn phải hỏi ngược lại (ví dụ đề hỏi mấy giờ đi mà chỉ nói 7h concert bắt đầu, hoặc viết câu vô nghĩa) -> trừ điểm Content.\n"
        f"   - Nếu thiếu kết thư hoặc thiếu ký tên người viết (kể cả chỉ ghi `Your friend.` không kèm tên) -> trừ 1 điểm Content (tối đa Content = 4).\n"
        f"   - Ở Part 1: Nếu đặt đủ 2 câu hỏi đúng chủ đề nhưng quên dấu `?` (dùng dấu `.`) thì KHÔNG trừ điểm Content. Ngược lại, câu mời xã giao `Would you like to my house?` / `OK for you?` KHÔNG tính là câu hỏi chủ đề.\n"
        f"   - Ở Part 2: Kiểm tra dòng `Đối chiếu chéo Part 1 của SV` — nếu sinh viên phớt lờ người gửi trên đề Part 2 mà tự đóng vai người nhận của Part 1 để viết thư hồi đáp cho chính bài Part 1 của mình -> Khóa cứng `0 - 1 - 1 - 1`!\n"
        f"2. NHẬN XÉT CHI TIẾT DÀNH CHO GIÁM KHẢO KIỂM CHỨNG (`c_reason`, `o_reason`, `g_reason`, `v_reason`):\n"
        f"   - Bắt buộc chỉ rõ đạt mấy ý / thiếu ý nào của đề và TRÍCH DẪN NGUYÊN VĂN cụm từ sai ngữ pháp/chính tả trong ngoặc kép `\"...\"`.\n"
        f"   - Có từ 2 lỗi cấu trúc ngữ pháp cơ bản trở lên -> chấm Grammar = 3 (hoặc 2 nếu lỗi nặng hệ thống), không chấm 4.\n"
        f"   - Mắc lỗi Comma splice (nối các câu đơn độc lập bằng dấu phẩy `,` không có liên từ) hoặc xuống dòng rời rạc từng câu đơn lẻ -> Organization tối đa = 3.\n"
        f"   - Chỉ dùng từ vựng cơ bản quen thuộc, không có cụm từ mở rộng -> Vocabulary tối đa = 3.\n"
        f"3. QUY TẮC KHÓA/PHẠT ĐIỂM:\n"
        f"   - Lạc đề hoàn toàn / học thuộc văn mẫu sai đề (Content = 0, bài >= 12 từ): khóa `0 - 1 - 1 - 1`.\n"
        f"   - Chép nguyên văn đề bài hoặc bài cụt < 12 từ: khóa `0 - 0 - 0 - 0`.\n\n"
        + "\n\n".join(section_blocks)
    )

    out_text = _call_gemini_interactions(
        api_key=api_key,
        input_data=prompt,
        schema_dict=KetQuaChamCumMaDe.model_json_schema(),
        preferred_model=model_name,
        timeout_sec=120,
        thinking_level=DEFAULT_THINKING_LEVEL,
    )
    cum_kq = KetQuaChamCumMaDe.model_validate_json(out_text)
    map_cum = {normalize_mssv(item.mssv): item for item in cum_kq.danh_sach}

    results: dict[str, KetQuaCham1PhanSV] = {}
    for mssv, sv_data, w_cnt in expected_items:
        item = map_cum.get(mssv)
        if item is not None:
            if w_cnt < 5:
                item.content = item.org = item.grammar = item.vocab = 0
            results[mssv] = item
        else:
            # Dự phòng nếu thiếu 1 SV trong cụm thì chấm lẻ SV đó
            _, kq_single = cham_1_sinh_vien(api_key, sv_data, rubric_text, model_name=model_name)
            if is_part1:
                results[mssv] = KetQuaCham1PhanSV(
                    mssv=mssv,
                    content=kq_single.p1_content,
                    org=kq_single.p1_org,
                    grammar=kq_single.p1_grammar,
                    vocab=kq_single.p1_vocab,
                    c_reason=kq_single.p1_c_reason,
                    o_reason=kq_single.p1_o_reason,
                    g_reason=kq_single.p1_g_reason,
                    v_reason=kq_single.p1_v_reason,
                    nhan_xet_ngan=kq_single.nhan_xet,
                )
            else:
                results[mssv] = KetQuaCham1PhanSV(
                    mssv=mssv,
                    content=kq_single.p2_content,
                    org=kq_single.p2_org,
                    grammar=kq_single.p2_grammar,
                    vocab=kq_single.p2_vocab,
                    c_reason=kq_single.p2_c_reason,
                    o_reason=kq_single.p2_o_reason,
                    g_reason=kq_single.p2_g_reason,
                    v_reason=kq_single.p2_v_reason,
                    nhan_xet_ngan=kq_single.nhan_xet,
                )
    return results


_LAST_COUNCIL_META_BY_MSSV: dict[str, dict] = {}


def _short_model_label(model_name: str) -> str:
    """Rút gọn tiền tố 'gemini-' để hiển thị nhãn model ngắn gọn, chuyên nghiệp trên thanh tiến trình."""
    m = (model_name or "").strip()
    if m.lower().startswith("gemini-"):
        return m[7:]
    return m or "3.8-flash"


def _select_borderline_cases_for_arbitration(
    text_tasks: list[tuple[int, str, str, dict]],
    raw_pass1_map: dict[str, KetQuaCham],
    calibrated_map: dict[str, KetQuaCham],
    max_cases: int = MAX_ARBITRATOR_CASES_PER_CLASS,
) -> list[dict]:
    """
    Lọc tự động các bài làm nhạy cảm / xung đột quy tắc để đưa vào Vòng 2 (Hội đồng Trọng tài Đa Model):
    1. Bài có sự xung đột giữa điểm của Giám khảo 1 (Pass 1 AI) và Bộ Quy tắc Khoa (_calibrate_scores).
    2. Bài nằm ở ranh giới nhạy cảm: nghi vấn lạc đề / thiếu ý nặng (Content <= 2) hoặc chênh lệch lớn giữa P1 và P2.
    3. Bài ở ngưỡng điểm cao (Grammar/Vocab >= 4) cần kiểm định chống chấm nương tay.
    """
    candidates: list[dict] = []
    fallback_candidates: list[dict] = []

    for row_idx, mssv, ten, sv_data in text_tasks:
        raw_kq = raw_pass1_map.get(mssv)
        cal_kq = calibrated_map.get(mssv)
        if raw_kq is None or cal_kq is None:
            continue

        reasons: list[str] = []
        priority = 0

        raw_tuple = (
            raw_kq.p1_content, raw_kq.p1_org, raw_kq.p1_grammar, raw_kq.p1_vocab,
            raw_kq.p2_content, raw_kq.p2_org, raw_kq.p2_grammar, raw_kq.p2_vocab,
        )
        cal_tuple = (
            cal_kq.p1_content, cal_kq.p1_org, cal_kq.p1_grammar, cal_kq.p1_vocab,
            cal_kq.p2_content, cal_kq.p2_org, cal_kq.p2_grammar, cal_kq.p2_vocab,
        )
        if raw_tuple != cal_tuple:
            priority += 50
            reasons.append(
                f"Lệch điểm giữa GK1 ({raw_kq.p1_content}-{raw_kq.p1_org}-{raw_kq.p1_grammar}-{raw_kq.p1_vocab} | "
                f"{raw_kq.p2_content}-{raw_kq.p2_org}-{raw_kq.p2_grammar}-{raw_kq.p2_vocab}) và Quy tắc Khoa "
                f"({cal_kq.p1_content}-{cal_kq.p1_org}-{cal_kq.p1_grammar}-{cal_kq.p1_vocab} | "
                f"{cal_kq.p2_content}-{cal_kq.p2_org}-{cal_kq.p2_grammar}-{cal_kq.p2_vocab})"
            )

        if 0 <= cal_kq.p1_content <= 2 or 0 <= cal_kq.p2_content <= 2:
            priority += 35
            reasons.append(f"Ranh giới lạc đề/thiếu ý (P1 C={cal_kq.p1_content}, P2 C={cal_kq.p2_content})")

        f1 = (cal_kq.p1_content + cal_kq.p1_org + cal_kq.p1_grammar + cal_kq.p1_vocab) / 4.0
        f2 = (cal_kq.p2_content + cal_kq.p2_org + cal_kq.p2_grammar + cal_kq.p2_vocab) / 4.0
        if abs(f1 - f2) >= 1.5:
            priority += 25
            reasons.append(f"Chênh lệch lớn giữa P1 ({f1:g}) và P2 ({f2:g})")

        if cal_kq.p1_grammar >= 4 or cal_kq.p1_vocab >= 4 or cal_kq.p2_grammar >= 4 or cal_kq.p2_vocab >= 4:
            priority += 15
            reasons.append("Kiểm định ngưỡng điểm G/V cao (4-5đ)")

        entry = {
            "row": row_idx,
            "mssv": mssv,
            "ten": ten,
            "sv_data": sv_data,
            "raw_kq": raw_kq,
            "cal_kq": cal_kq,
            "priority": priority,
            "conflict_reasons": reasons if reasons else ["Rà soát chéo ngẫu nhiên chuẩn Hội đồng 2 vòng"],
        }
        if priority > 0:
            candidates.append(entry)
        else:
            fallback_candidates.append(entry)

    candidates.sort(key=lambda x: (-x["priority"], x["row"]))
    selected = candidates[:max_cases]
    if not selected and fallback_candidates:
        selected = fallback_candidates[: min(3, max_cases, len(fallback_candidates))]
    return selected


def cham_vong_2_trong_tai_phan_bien(
    api_key: str,
    candidate_entries: list[dict],
    rubric_text: str,
    pass1_model: str = DEFAULT_MODEL,
    arbitrator_model: str = DEFAULT_ARBITRATOR_MODEL,
) -> tuple[dict[str, KetQuaTrongTaiSV], str]:
    """
    Vòng 2 (Hội đồng Trọng tài Độc lập bằng Model thứ 2):
    Nhận danh sách 3-6 bài nhạy cảm/xung đột trong lớp (chỉ tốn đúng 1 lượt gọi API/lớp),
    đối chiếu độc lập với Đề gốc + Bài mẫu Gold Standard + Điểm đề xuất của GK1 + Cảnh báo Quy tắc Khoa để chốt điểm cuối cùng.
    """
    if not candidate_entries:
        return {}, ""

    case_blocks: list[str] = []
    for idx, item in enumerate(candidate_entries, start=1):
        mssv = item["mssv"]
        ten = item["ten"]
        sv_data = item["sv_data"]
        raw_kq: KetQuaCham = item["raw_kq"]
        cal_kq: KetQuaCham = item["cal_kq"]
        reasons_txt = "; ".join(item.get("conflict_reasons") or [])

        p1 = sv_data.get("p1") or {}
        p2 = sv_data.get("p2") or {}
        c1 = (p1.get("code") or "N/A").strip()
        c2 = (p2.get("code") or "N/A").strip()
        pr1 = clean_prompt_text(p1.get("prompt") or "") or "[Không có đề P1]"
        pr2 = clean_prompt_text(p2.get("prompt") or "") or "[Không có đề P2]"
        gold1 = _format_gold_examples_for_code(c1, is_part1=True, max_examples=1)
        gold2 = _format_gold_examples_for_code(c2, is_part1=False, max_examples=1)
        es1 = _format_essay_with_line_numbers(p1)
        es2 = _format_essay_with_line_numbers(p2)
        p1_cross = _extract_p1_cross_ref_note(sv_data)

        case_blocks.append(
            f"====================================================================\n"
            f"[HỒ SƠ THẨM ĐỊNH #{idx} — MSSV: {mssv} | HỌ TÊN: {ten}]\n"
            f"• Lý do đưa ra Hội đồng Trọng tài: {reasons_txt}\n"
            f"• Điểm Giám khảo 1 ({pass1_model}): P1=({raw_kq.p1_content}-{raw_kq.p1_org}-{raw_kq.p1_grammar}-{raw_kq.p1_vocab}) | "
            f"P2=({raw_kq.p2_content}-{raw_kq.p2_org}-{raw_kq.p2_grammar}-{raw_kq.p2_vocab})\n"
            f"• Điểm Hậu kiểm Quy tắc Khoa (Python Rule Engine): P1=({cal_kq.p1_content}-{cal_kq.p1_org}-{cal_kq.p1_grammar}-{cal_kq.p1_vocab}) | "
            f"P2=({cal_kq.p2_content}-{cal_kq.p2_org}-{cal_kq.p2_grammar}-{cal_kq.p2_vocab})\n"
            f"• Nhận xét tạm thời: P1[C:{cal_kq.p1_c_reason} | O:{cal_kq.p1_o_reason} | G:{cal_kq.p1_g_reason} | V:{cal_kq.p1_v_reason}] "
            f"// P2[C:{cal_kq.p2_c_reason} | O:{cal_kq.p2_o_reason} | G:{cal_kq.p2_g_reason} | V:{cal_kq.p2_v_reason}]\n\n"
            f"--- [PART 1 · Mã đề {c1}] ---\n"
            f"Đề gốc P1:\n{pr1}\n"
            f"Bài mẫu chuẩn GV (Gold Standard {c1}):\n{gold1}\n"
            f"Bài làm P1 của SV {mssv}:\n{es1}\n\n"
            f"--- [PART 2 · Mã đề {c2}] ---\n"
            f"Đề gốc P2:\n{pr2}\n"
            f"Bài mẫu chuẩn GV (Gold Standard {c2}):\n{gold2}\n"
            f"{p1_cross}\n"
            f"Bài làm P2 của SV {mssv}:\n{es2}"
        )

    prompt = (
        f"Bạn là CHỦ TỊCH HỘI ĐỒNG CHẤM THI PHÚC KHẢO (Chief Examiner / Arbitrator - Vòng 2) môn Anh Văn Căn Bản (Writing A1) tại trường Đại học Công nghiệp TP.HCM (IUH).\n"
        f"Nhiệm vụ: Phản biện chéo và chốt điểm cuối cùng cho {len(candidate_entries)} hồ sơ bài thi nhạy cảm/xung đột giữa Giám khảo 1 ({pass1_model}) và Hệ thống Hậu kiểm Quy tắc Khoa.\n\n"
        f"{rubric_text}\n\n"
        f"NGUYÊN TẮC PHẢN BIỆN CỦA HỘI ĐỒNG TRỌNG TÀI VÒNG 2:\n"
        f"1. Đọc độc lập bài làm P1 & P2 của từng sinh viên, đối chiếu trực tiếp với Đề bài gốc và Bài mẫu chuẩn Gold Standard cùng mã đề.\n"
        f"2. Kiểm tra nơi xảy ra bất đồng giữa Giám khảo 1 và Quy tắc Khoa:\n"
        f"   - Nếu sinh viên thực sự thiếu kết thư/ký tên, hoặc xuống dòng rời rạc từng câu, hoặc mắc lỗi Comma splice, hoặc chỉ dùng từ cơ bản -> Bắt buộc giữ mức trừ điểm nghiêm khắc của Quy tắc Khoa.\n"
        f"   - Nếu bài làm lạc đề hoàn toàn hoặc ở Part 2 tự đóng vai người nhận của Part 1 để trả lời thư P1 -> Khóa cứng `0 - 1 - 1 - 1`.\n"
        f"   - Ghi rõ quyết định vào trường `verdict` (ví dụ: `Đồng thuận Quy tắc Khoa: trừ điểm thiếu ký tên` hoặc `Đồng thuận GK1: đủ 3 ý chuẩn`) và trích dẫn nguyên văn lỗi trong `\"...\"` ở 8 cột lý do.\n\n"
        + "\n\n".join(case_blocks)
    )

    out_text = _call_gemini_interactions(
        api_key=api_key,
        input_data=prompt,
        schema_dict=KetQuaHoiDongTrongTai.model_json_schema(),
        preferred_model=arbitrator_model or DEFAULT_ARBITRATOR_MODEL,
        timeout_sec=120,
        thinking_level=DEFAULT_THINKING_LEVEL,
    )
    arb_kq = KetQuaHoiDongTrongTai.model_validate_json(out_text)
    arb_map = {normalize_mssv(it.mssv): it for it in arb_kq.danh_sach}
    return arb_map, (arb_kq.nhan_xet_hoi_dong or "").strip()


def _grade_tasks_by_prompt_clusters(
    api_key: str,
    tasks: list[tuple[int, str, str, dict]],
    rubric_text: str,
    model_name: str = DEFAULT_MODEL,
    stop_event: threading.Event | None = None,
    enable_pass2_arbitrator: bool = False,
    arbitrator_model: str = DEFAULT_ARBITRATOR_MODEL,
    stage_cb=None,
) -> tuple[dict[str, KetQuaCham], dict[str, Exception], int]:
    """
    Điều phối chấm cả lớp theo Quy trình Hội đồng Đa Tầng (Multi-Stage & Multi-Model AI Council):
    - Bước 1: Phân tích & Gom ngầm 20 Mã đề (`P1.L001..L010`, `P2.L001..L010`) + Nạp Bài mẫu Gold Standard (0s, giữ nguyên thứ tự SV).
    - Bước 2: Giám khảo 1 (`model_name`, mặc định `gemini-3.8-flash` High Thinking) chấm 4 Cụm Mã Đề (2 cụm P1 + 2 cụm P2).
    - Bước 3: Hậu kiểm Quy tắc Khoa (`_calibrate_scores` — Python Rule Engine) rà soát 12 quy tắc cứng.
    - Bước 4 (Tùy chọn / Mặc định bật trên GUI): Trọng tài Vòng 2 (`arbitrator_model`, mặc định `gemini-3.7-flash`) phản biện chéo các ca nhạy cảm/lệch điểm trong 1 lượt gọi API.
    - Bước 5: Tổng hợp dẫn chứng nguyên văn `"..."` và hoàn trả về đúng dòng Excel gốc.
    """
    global _ACTIVE_THOUGHT_CALLBACK
    results_by_mssv: dict[str, KetQuaCham] = {}
    errors_by_mssv: dict[str, Exception] = {}
    text_tasks: list[tuple[int, str, str, dict]] = []
    api_calls_count = 0
    _LAST_COUNCIL_META_BY_MSSV.clear()

    def _notify_stage(step_idx: int, state: str, badge_txt: str, detail_msg: str):
        if stage_cb is not None:
            try:
                stage_cb(step_idx, state, badge_txt, detail_msg)
            except Exception:
                pass

    prev_thought_cb = _ACTIVE_THOUGHT_CALLBACK

    def _on_live_thought(m_name: str, summary_txt: str):
        short_s = re.sub(r"\s+", " ", summary_txt).strip()
        if len(short_s) > 260:
            short_s = short_s[:257] + "..."
        _notify_stage(0, "thought", m_name, f"[Suy luận nội bộ · {m_name}]: {short_s}")
        if prev_thought_cb is not None:
            try:
                prev_thought_cb(m_name, summary_txt)
            except Exception:
                pass

    _ACTIVE_THOUGHT_CALLBACK = _on_live_thought
    try:
        for row_idx, mssv, ten, sv_data in tasks:
            if stop_event and stop_event.is_set():
                break
            p1 = sv_data.get("p1") or {}
            p2 = sv_data.get("p2") or {}
            w1 = count_words((p1.get("essay") or "").strip())
            w2 = count_words((p2.get("essay") or "").strip())

            if w1 == 0 and w2 == 0:
                kq_empty = KetQuaCham(
                    p1_content=0, p1_org=0, p1_grammar=0, p1_vocab=0,
                    p2_content=0, p2_org=0, p2_grammar=0, p2_vocab=0,
                    nhan_xet="Sinh viên bỏ trắng cả 2 bài (0 từ).",
                )
                results_by_mssv[mssv] = ensure_cogv_reasons(sv_data, kq_empty)
                _LAST_COUNCIL_META_BY_MSSV[mssv] = {
                    "stage": "BO_TRONG",
                    "badge": "Bỏ trống (0 từ)",
                    "verdict": "Sinh viên bỏ trắng cả 2 bài.",
                }
            elif (p1.get("images") or []) or (p2.get("images") or []):
                try:
                    api_calls_count += 1
                    _, kq_img = cham_1_sinh_vien(api_key, sv_data, rubric_text, model_name=model_name)
                    results_by_mssv[mssv] = kq_img
                    _LAST_COUNCIL_META_BY_MSSV[mssv] = {
                        "stage": "1_VONG_ANH",
                        "badge": f"Đã chấm Multimodal ({model_name})",
                        "verdict": "Chấm trực tiếp kèm hình ảnh đề bài.",
                    }
                except Exception as err:
                    errors_by_mssv[mssv] = err
            else:
                text_tasks.append((row_idx, mssv, ten, sv_data))

        if not text_tasks or (stop_event and stop_event.is_set()):
            return results_by_mssv, errors_by_mssv, api_calls_count

        # ==================== BƯỚC 1: PHÂN CỤM 20 MÃ ĐỀ & NẠP GOLD STANDARD ====================
        _notify_stage(
            1,
            "active",
            "Đang gom mã đề...",
            f"[Bước 1] Đang phân loại ngầm {len(text_tasks)} SV theo 20 mã đề chuẩn (P1.L001..L010 & P2.L001..L010) và nạp bài mẫu Gold Standard...",
        )
        p1_clusters = _build_prompt_clusters_for_part(text_tasks, "p1", max_clusters=MAX_PROMPT_CLUSTERS_PER_PART)
        p2_clusters = _build_prompt_clusters_for_part(text_tasks, "p2", max_clusters=MAX_PROMPT_CLUSTERS_PER_PART)
        p1_code_cnt = sum(len(c) for c in p1_clusters)
        p2_code_cnt = sum(len(c) for c in p2_clusters)
        _notify_stage(
            1,
            "done",
            f"Đã gom {p1_code_cnt} mã P1 · {p2_code_cnt} mã P2",
            f"[Bước 1 · Hoàn tất] Đã chia ngầm {len(text_tasks)} SV vào {len(p1_clusters)} cụm P1 ({p1_code_cnt} mã đề) & {len(p2_clusters)} cụm P2 ({p2_code_cnt} mã đề) kèm 20/20 bài mẫu Gold Standard (Bảo toàn 100% thứ tự dòng SV).",
        )

        # ==================== BƯỚC 2: GIÁM KHẢO 1 CHẤM CỤM MÃ ĐỀ (HIGH THINKING) ====================
        cluster_jobs: list[tuple[str, int, list[tuple[str, list[tuple[int, str, str, dict]]]]]] = []
        for idx_c, c_grp in enumerate(p1_clusters, start=1):
            cluster_jobs.append(("p1", idx_c, c_grp))
        for idx_c, c_grp in enumerate(p2_clusters, start=1):
            cluster_jobs.append(("p2", idx_c, c_grp))

        total_clusters = len(cluster_jobs)
        _notify_stage(
            2,
            "active",
            f"GK1 ({model_name}) · 0/{total_clusters} cụm",
            f"[Bước 2 · Giám khảo 1 ({model_name} · High Thinking)] Bắt đầu chấm {total_clusters} cụm mã đề song song (Đối chiếu Đề gốc + Gold Standard + Chéo P1→P2)...",
        )

        p1_res_map: dict[str, KetQuaCham1PhanSV] = {}
        p2_res_map: dict[str, KetQuaCham1PhanSV] = {}
        p1_err_map: dict[str, Exception] = {}
        p2_err_map: dict[str, Exception] = {}

        done_clusters = 0
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_job = {}
            for j_idx, (pk, c_num, c_grp) in enumerate(cluster_jobs):
                if stop_event and stop_event.is_set():
                    break
                api_calls_count += 1
                codes_str = ", ".join(code for code, _ in c_grp[:4])
                sv_cnt = sum(len(st_list) for _, st_list in c_grp)
                _notify_stage(
                    2,
                    "active",
                    f"GK1 ({model_name}) · Đang suy luận...",
                    f"[Bước 2 · GK1 {model_name}] Đang gửi Cụm {c_num} {pk.upper()} ({sv_cnt} bài · Mã đề: {codes_str})...",
                )
                fut = executor.submit(cham_1_cum_ma_de, api_key, c_grp, pk, rubric_text, model_name)
                future_to_job[fut] = (pk, c_num, c_grp)
                if j_idx < len(cluster_jobs) - 1:
                    time.sleep(0.4)

            for future in as_completed(future_to_job):
                if stop_event and stop_event.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                pk, c_num, c_grp = future_to_job[future]
                done_clusters += 1
                try:
                    part_res = future.result()
                    if pk == "p1":
                        p1_res_map.update(part_res)
                    else:
                        p2_res_map.update(part_res)
                    _notify_stage(
                        2,
                        "active" if done_clusters < total_clusters else "done",
                        f"GK1 hoàn tất {done_clusters}/{total_clusters} cụm",
                        f"[Bước 2 · GK1 {model_name}] Đã chấm xong Cụm {c_num} {pk.upper()} ({len(part_res)} SV) — Tiến độ cụm: {done_clusters}/{total_clusters}.",
                    )
                except Exception as err:
                    for _, st_list in c_grp:
                        for _, mssv, _, _ in st_list:
                            if pk == "p1":
                                p1_err_map[mssv] = err
                            else:
                                p2_err_map[mssv] = err

        # ==================== BƯỚC 3: HỆ THỐNG HẬU KIỂM QUY TẮC KHOA (PYTHON RULE ENGINE) ====================
        _notify_stage(
            3,
            "active",
            "Đang quét 12 quy tắc Khoa...",
            "[Bước 3 · Python Rule Engine] Đang kiểm tra chữ ký cuối thư, lỗi Comma splice, xuống dòng rời rạc, câu hỏi chủ đề P1 và lộn vai P1→P2...",
        )
        raw_pass1_map: dict[str, KetQuaCham] = {}
        calibrated_map: dict[str, KetQuaCham] = {}
        rule_adjusted_cnt = 0

        for _, mssv, _, sv_data in text_tasks:
            if stop_event and stop_event.is_set():
                break
            p1 = sv_data.get("p1") or {}
            p2 = sv_data.get("p2") or {}
            w1 = count_words((p1.get("essay") or "").strip())
            w2 = count_words((p2.get("essay") or "").strip())

            if w1 > 0 and mssv in p1_err_map:
                errors_by_mssv[mssv] = p1_err_map[mssv]
                continue
            if w2 > 0 and mssv in p2_err_map:
                errors_by_mssv[mssv] = p2_err_map[mssv]
                continue

            r1 = p1_res_map.get(mssv) if w1 > 0 else KetQuaCham1PhanSV(
                mssv=mssv, content=0, org=0, grammar=0, vocab=0,
                c_reason="Bỏ trống Part 1 (0 từ).", o_reason="Bỏ trống Part 1 (0 từ).",
                g_reason="Bỏ trống Part 1 (0 từ).", v_reason="Bỏ trống Part 1 (0 từ).",
            )
            r2 = p2_res_map.get(mssv) if w2 > 0 else KetQuaCham1PhanSV(
                mssv=mssv, content=0, org=0, grammar=0, vocab=0,
                c_reason="Bỏ trống Part 2 (0 từ).", o_reason="Bỏ trống Part 2 (0 từ).",
                g_reason="Bỏ trống Part 2 (0 từ).", v_reason="Bỏ trống Part 2 (0 từ).",
            )
            if r1 is None or r2 is None:
                errors_by_mssv[mssv] = RuntimeError("Không nhận được kết quả từ cụm mã đề AI.")
                continue

            nx_parts = [s.strip() for s in (r1.nhan_xet_ngan, r2.nhan_xet_ngan) if s and s.strip()]
            raw_kq = KetQuaCham(
                nhan_xet=" | ".join(nx_parts) if nx_parts else "Đã chấm theo cụm mã đề chuẩn.",
                p1_content=0 if w1 < 5 else r1.content,
                p1_org=0 if w1 < 5 else r1.org,
                p1_grammar=0 if w1 < 5 else r1.grammar,
                p1_vocab=0 if w1 < 5 else r1.vocab,
                p1_c_reason=r1.c_reason,
                p1_o_reason=r1.o_reason,
                p1_g_reason=r1.g_reason,
                p1_v_reason=r1.v_reason,
                p2_content=0 if w2 < 5 else r2.content,
                p2_org=0 if w2 < 5 else r2.org,
                p2_grammar=0 if w2 < 5 else r2.grammar,
                p2_vocab=0 if w2 < 5 else r2.vocab,
                p2_c_reason=r2.c_reason,
                p2_o_reason=r2.o_reason,
                p2_g_reason=r2.g_reason,
                p2_v_reason=r2.v_reason,
            )
            raw_pass1_map[mssv] = raw_kq.model_copy(deep=True)
            cal_kq = _calibrate_scores(sv_data, raw_kq.model_copy(deep=True))
            cal_kq = ensure_cogv_reasons(sv_data, cal_kq)
            calibrated_map[mssv] = cal_kq
            results_by_mssv[mssv] = cal_kq

            if (
                raw_kq.p1_content != cal_kq.p1_content or raw_kq.p1_org != cal_kq.p1_org
                or raw_kq.p1_grammar != cal_kq.p1_grammar or raw_kq.p1_vocab != cal_kq.p1_vocab
                or raw_kq.p2_content != cal_kq.p2_content or raw_kq.p2_org != cal_kq.p2_org
                or raw_kq.p2_grammar != cal_kq.p2_grammar or raw_kq.p2_vocab != cal_kq.p2_vocab
            ):
                rule_adjusted_cnt += 1
                _LAST_COUNCIL_META_BY_MSSV[mssv] = {
                    "stage": "1_VONG_HAU_KIEM",
                    "badge": f"Chuẩn hội đồng ({_short_model_label(model_name)} + Rule Engine)",
                    "verdict": "Đã hiệu chỉnh khung điểm theo Quy tắc Khoa IUH.",
                }
            else:
                _LAST_COUNCIL_META_BY_MSSV[mssv] = {
                    "stage": "1_VONG_DONG_THUAN",
                    "badge": f"Chuẩn hội đồng ({_short_model_label(model_name)} + Rule Engine)",
                    "verdict": "Đồng thuận 100% giữa GK1 và Quy tắc Khoa.",
                }

        borderline_cases = _select_borderline_cases_for_arbitration(
            text_tasks, raw_pass1_map, calibrated_map, max_cases=MAX_ARBITRATOR_CASES_PER_CLASS
        )
        _notify_stage(
            3,
            "done",
            f"Đã hậu kiểm · {rule_adjusted_cnt} ca siết điểm",
            f"[Bước 3 · Hoàn tất] Đã rà soát {len(calibrated_map)} SV qua Bộ Quy tắc Khoa ({rule_adjusted_cnt} bài được hiệu chỉnh khung điểm, phát hiện {len(borderline_cases)} hồ sơ nhạy cảm).",
        )

        # ==================== BƯỚC 4: HỘI ĐỒNG TRỌNG TÀI VÒNG 2 (ĐA MÔ HÌNH PHẢN BIỆN CHÉO) ====================
        if enable_pass2_arbitrator and borderline_cases and not (stop_event and stop_event.is_set()):
            arb_model = (arbitrator_model or DEFAULT_ARBITRATOR_MODEL).strip()
            sv_names_preview = ", ".join(f"{c['mssv']} ({c['ten']})" for c in borderline_cases[:4])
            _notify_stage(
                4,
                "active",
                f"Trọng tài ({arb_model}) · {len(borderline_cases)} bài",
                f"[Bước 4 · Trọng tài Vòng 2 ({arb_model} · High Thinking)] Đang phúc khảo độc lập {len(borderline_cases)} hồ sơ nhạy cảm/lệch điểm: {sv_names_preview}...",
            )
            try:
                api_calls_count += 1
                arb_map, council_summary = cham_vong_2_trong_tai_phan_bien(
                    api_key=api_key,
                    candidate_entries=borderline_cases,
                    rubric_text=rubric_text,
                    pass1_model=model_name,
                    arbitrator_model=arb_model,
                )
                arbitrated_ok = 0
                for cand in borderline_cases:
                    m_id = cand["mssv"]
                    sv_d = cand["sv_data"]
                    arb_item = arb_map.get(m_id)
                    if arb_item is None:
                        continue
                    w1 = count_words(((sv_d.get("p1") or {}).get("essay") or "").strip())
                    w2 = count_words(((sv_d.get("p2") or {}).get("essay") or "").strip())
                    verdict_str = (arb_item.verdict or "Đã thẩm định 2 vòng").strip()
                    kq_arb = KetQuaCham(
                        nhan_xet=f"[Hội đồng 2 vòng · {arb_model}]: {verdict_str}",
                        p1_content=0 if w1 < 5 else arb_item.p1_content,
                        p1_org=0 if w1 < 5 else arb_item.p1_org,
                        p1_grammar=0 if w1 < 5 else arb_item.p1_grammar,
                        p1_vocab=0 if w1 < 5 else arb_item.p1_vocab,
                        p1_c_reason=arb_item.p1_c_reason,
                        p1_o_reason=arb_item.p1_o_reason,
                        p1_g_reason=arb_item.p1_g_reason,
                        p1_v_reason=arb_item.p1_v_reason,
                        p2_content=0 if w2 < 5 else arb_item.p2_content,
                        p2_org=0 if w2 < 5 else arb_item.p2_org,
                        p2_grammar=0 if w2 < 5 else arb_item.p2_grammar,
                        p2_vocab=0 if w2 < 5 else arb_item.p2_vocab,
                        p2_c_reason=arb_item.p2_c_reason,
                        p2_o_reason=arb_item.p2_o_reason,
                        p2_g_reason=arb_item.p2_g_reason,
                        p2_v_reason=arb_item.p2_v_reason,
                    )
                    kq_arb = _calibrate_scores(sv_d, kq_arb)
                    kq_arb = ensure_cogv_reasons(sv_d, kq_arb)
                    results_by_mssv[m_id] = kq_arb
                    _LAST_COUNCIL_META_BY_MSSV[m_id] = {
                        "stage": "2_VONG_TRONG_TAI",
                        "badge": f"Thẩm định 2 vòng ({_short_model_label(model_name)} ⇄ {_short_model_label(arb_model)})",
                        "verdict": verdict_str,
                    }
                    arbitrated_ok += 1

                summary_note = f" — Kết luận: {council_summary}" if council_summary else ""
                _notify_stage(
                    4,
                    "done",
                    f"Đã phúc khảo {arbitrated_ok}/{len(borderline_cases)} bài",
                    f"[Bước 4 · Hoàn tất] Trọng tài Vòng 2 ({arb_model}) đã thẩm định độc lập {arbitrated_ok}/{len(borderline_cases)} bài nhạy cảm{summary_note}",
                )
            except Exception as arb_err:
                _notify_stage(
                    4,
                    "done",
                    "Giữ điểm Hậu kiểm Vòng 1",
                    f"[Bước 4 · Dự phòng an toàn] Vòng 2 ({arb_model}) bỏ qua do giới hạn kết nối ({arb_err}) -> Giữ nguyên kết quả chuẩn Vòng 1 + Rule Engine.",
                )
        else:
            _notify_stage(
                4,
                "done",
                "Đồng thuận Vòng 1 + Rule Engine",
                "[Bước 4] Chế độ 1 vòng tốc độ cao (hoặc không có ca xung đột) — Đã chốt điểm qua GK1 + Rule Engine.",
            )

        # ==================== BƯỚC 5: ĐỒNG BỘ DẪN CHỨNG & BẢO TOÀN THỨ TỰ DÒNG EXCEL ====================
        _notify_stage(
            5,
            "done",
            f"Hoàn tất {len(results_by_mssv)} SV · {api_calls_count} lượt API",
            f"[Bước 5 · Hoàn tất] Đã gắn dẫn chứng nguyên văn \"...\" cho {len(results_by_mssv)} SV và hoàn trả về đúng 100% thứ tự dòng Excel gốc (Tổng cộng: {api_calls_count} lượt gọi API).",
        )
        return results_by_mssv, errors_by_mssv, api_calls_count
    finally:
        _ACTIVE_THOUGHT_CALLBACK = prev_thought_cb


def extract_explicit_iuh_class_code(path_str: str) -> str:
    """Trích xuất mã lớp IUH tường minh (VD: 2607AVCB005) từ đường dẫn file; trả về chuỗi rỗng nếu không có mã chuẩn."""
    if not path_str:
        return ""
    for text in (os.path.basename(path_str), os.path.basename(os.path.dirname(os.path.abspath(path_str)))):
        m = re.search(r"(\d{4}\s*[A-Z]{3,8}\.?\s*\d{2,4})", text, flags=re.IGNORECASE)
        if m:
            return re.sub(r"[\s.]+", "", m.group(1)).upper()
    return ""


def validate_class_files_not_mixed(json_path: str, excel_path: str) -> None:
    """Chốt chặn bảo vệ file gốc trên máy tính: Ngăn chặn ghép nhầm file JSON của lớp A với file Excel của lớp B."""
    if not json_path or not excel_path:
        return
    j_code = extract_explicit_iuh_class_code(json_path)
    e_code = extract_explicit_iuh_class_code(excel_path)
    if j_code and e_code and j_code != e_code:
        raise ValueError(
            f"CẢNH BÁO LẪN LỘN FILE GIỮA 2 LỚP KHÁC NHAU:\n"
            f"• File JSON thuộc lớp: {j_code} ({os.path.basename(json_path)})\n"
            f"• File Excel thuộc lớp: {e_code} ({os.path.basename(excel_path)})\n\n"
            f"Hệ thống đã tự động chặn ghép nối để bảo vệ tuyệt đối file gốc trên máy tính.\n"
            f"Hãy chọn đúng file cùng lớp hoặc nhấn '+ Mở lớp' trên thanh tiêu đề để mở lớp {e_code} ở tab riêng!"
        )


def resolve_or_recover_class_files(code: str = "", json_path: str = "", excel_path: str = "") -> tuple[str, str]:
    """Tự động khôi phục đường dẫn file JSON / Excel của lớp nếu 1 trong 2 đường dẫn bị thiếu hoặc từng trỏ vào thư mục tạm."""
    j_p = (json_path or "").strip()
    e_p = (excel_path or "").strip()
    detected_code = (
        extract_explicit_iuh_class_code(e_p)
        or extract_explicit_iuh_class_code(j_p)
        or (code.strip().upper() if re.match(r"^\d{4}[A-Z]{3,8}\d{2,4}$", (code or "").strip(), re.IGNORECASE) else "")
    )

    search_dirs = []
    if e_p and os.path.exists(os.path.dirname(os.path.abspath(e_p))):
        search_dirs.append(os.path.dirname(os.path.abspath(e_p)))
    if j_p and os.path.exists(os.path.dirname(os.path.abspath(j_p))):
        d_j = os.path.dirname(os.path.abspath(j_p))
        if d_j not in search_dirs:
            search_dirs.append(d_j)
    app_dir = get_app_dir()
    if detected_code:
        cls_subdir = os.path.join(app_dir, detected_code)
        if os.path.isdir(cls_subdir) and cls_subdir not in search_dirs:
            search_dirs.append(cls_subdir)
    if os.path.isdir(app_dir) and app_dir not in search_dirs:
        search_dirs.append(app_dir)

    # 1. Khôi phục json_path nếu bị mất hoặc không tồn tại
    if not j_p or not os.path.exists(j_p):
        for s_dir in search_dirs:
            try:
                files = os.listdir(s_dir)
            except Exception:
                continue
            raw_jsons = [
                os.path.join(s_dir, f)
                for f in files
                if f.lower().endswith(".json")
                and not f.lower().endswith("_review.json")
                and f.lower() != "config.json"
            ]
            if detected_code:
                matched_j = [p for p in raw_jsons if detected_code.lower() in os.path.basename(p).lower()]
                if matched_j:
                    j_p = matched_j[0]
                    break
            elif len(raw_jsons) == 1 and s_dir != app_dir:
                j_p = raw_jsons[0]
                break

    # 2. Khôi phục excel_path nếu bị mất hoặc không tồn tại
    if not e_p or not os.path.exists(e_p):
        for s_dir in search_dirs:
            try:
                files = os.listdir(s_dir)
            except Exception:
                continue
            xlsx_list = [
                os.path.join(s_dir, f)
                for f in files
                if f.lower().endswith(".xlsx") and not f.startswith("~$")
            ]
            if detected_code:
                matched_x = [p for p in xlsx_list if detected_code.lower() in os.path.basename(p).lower()]
                if matched_x:
                    matched_x.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
                    e_p = matched_x[0]
                    break
            elif len(xlsx_list) == 1 and s_dir != app_dir:
                e_p = xlsx_list[0]
                break

    return j_p, e_p


def get_review_session_path(json_path: str = "", excel_path: str = "", class_code: str = "") -> str:
    """Đường dẫn file lưu phiên kiểm duyệt trung gian (<tên_json>_review.json), tách biệt độc lập cho từng lớp."""
    if json_path:
        abs_j = os.path.abspath(json_path)
        if abs_j.lower().endswith("_review.json"):
            return abs_j
        base, _ = os.path.splitext(abs_j)
        return f"{base}_review.json"
    if excel_path:
        folder = os.path.dirname(os.path.abspath(excel_path))
        code = class_code or extract_explicit_iuh_class_code(excel_path) or "LOP"
        cand = os.path.join(folder, f"bai_lam_iuh_{code}_review.json")
        if os.path.exists(cand):
            return cand
        if os.path.isdir(folder):
            for f in os.listdir(folder):
                if f.lower().endswith("_review.json") and (not code or code.lower() in f.lower()):
                    return os.path.join(folder, f)
        return cand
    fallback_dir = os.path.dirname(os.path.abspath(get_config_path()))
    return os.path.join(fallback_dir, f"bai_lam_iuh_{class_code or 'session'}_review.json")


def save_review_session(json_path: str, excel_path: str, items: list[dict], class_code: str = "") -> str:
    """Lưu danh sách bài chấm & trạng thái tick kiểm duyệt của giảng viên ra file _review.json (Có khóa bảo vệ chống ghi đè nhầm lớp)."""
    validate_class_files_not_mixed(json_path, excel_path)
    code = (
        extract_explicit_iuh_class_code(class_code)
        or extract_explicit_iuh_class_code(json_path)
        or extract_explicit_iuh_class_code(excel_path)
        or (class_code or "").strip()
    )
    review_path = get_review_session_path(json_path, excel_path, class_code=code)

    new_mssvs = {normalize_mssv(it.get("mssv")) for it in items if normalize_mssv(it.get("mssv"))}
    if os.path.exists(review_path) and new_mssvs:
        try:
            with open(review_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
        except Exception:
            old_data = None
        old_items = old_data.get("items") if isinstance(old_data, dict) else None
        if isinstance(old_items, list) and old_items:
            old_mssvs = {normalize_mssv(x.get("mssv")) for x in old_items if normalize_mssv(x.get("mssv"))}
            if old_mssvs and len(new_mssvs & old_mssvs) == 0:
                raise ValueError(
                    f"CẢNH BÁO BẢO VỆ BỘ NHỚ LỚP GỐC:\n"
                    f"File phiên '{os.path.basename(review_path)}' đang lưu danh sách sinh viên của lớp khác "
                    f"(0 MSSV trùng khớp). Hệ thống đã chặn ghi đè để tránh mất dữ liệu đã chấm!"
                )

    payload = {
        "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "class_code": code,
        "json_path": os.path.abspath(json_path) if json_path else "",
        "excel_path": os.path.abspath(excel_path) if excel_path else "",
        "items": items,
    }
    with open(review_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return review_path


def load_review_session(json_path: str = "", excel_path: str = "", class_code: str = "") -> dict | None:
    """Đọc lại phiên kiểm duyệt đã lưu từ file _review.json (nếu có)."""
    if not json_path and not excel_path and not class_code:
        return None
    review_path = get_review_session_path(json_path, excel_path, class_code=class_code)
    if not os.path.exists(review_path) and excel_path:
        review_path = get_review_session_path("", excel_path, class_code=class_code)
    if not os.path.exists(review_path):
        return None
    try:
        with open(review_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("items"), list):
            return data
    except Exception:
        return None
    return None


load_review_session_if_exists = load_review_session


def _detect_student_warning_tag(sv_data: dict, kq: KetQuaCham) -> str:
    """Xác định nhãn cảnh báo ngắn (Lạc đề, Chép đề, Lan man, Bỏ trống) hiển thị trên danh sách sinh viên."""
    p1 = sv_data.get("p1") or {}
    p2 = sv_data.get("p2") or {}
    w1 = count_words(p1.get("essay", ""))
    w2 = count_words(p2.get("essay", ""))
    if w1 == 0 and w2 == 0:
        return "Bỏ trống"
    if (kq.p1_content == 0 and w1 >= 12) or (kq.p2_content == 0 and w2 >= 12):
        if "chép lại" in (kq.nhan_xet or "").lower():
            return "Chép đề"
        return "Lạc đề"
    if w1 >= 95 or w2 >= 95:
        return "Lan man"
    if w1 < 12 or w2 < 12:
        return "Bài cụt"
    return ""


def _make_review_item(row_idx: int, mssv: str, ten: str, sv_data: dict | None, kq: KetQuaCham | None, checked: bool = False) -> dict:
    """Tạo cấu trúc dữ liệu chuẩn của 1 sinh viên cho Tab Kiểm duyệt."""
    if sv_data is None:
        return {
            "row": row_idx,
            "mssv": mssv,
            "ten": ten,
            "checked": checked,
            "absent": True,
            "graded": False,
            "note": "Vắng thi",
            "p1": {
                "code": "VẮNG THI",
                "prompt": "",
                "images": [],
                "essay": "[VẮNG THI - Không có bài trên hệ thống Moodle]",
                "essay_lines": [],
                "word_count": 0,
                "line_count": 0,
                "format_type": "TRONG",
                "c": 0, "o": 0, "g": 0, "v": 0,
                "rc": "Vắng thi trên Moodle.",
                "ro": "Vắng thi trên Moodle.",
                "rg": "Vắng thi trên Moodle.",
                "rv": "Vắng thi trên Moodle.",
            },
            "p2": {
                "code": "VẮNG THI",
                "prompt": "",
                "images": [],
                "essay": "[VẮNG THI - Không có bài trên hệ thống Moodle]",
                "essay_lines": [],
                "word_count": 0,
                "line_count": 0,
                "format_type": "TRONG",
                "c": 0, "o": 0, "g": 0, "v": 0,
                "rc": "Vắng thi trên Moodle.",
                "ro": "Vắng thi trên Moodle.",
                "rg": "Vắng thi trên Moodle.",
                "rv": "Vắng thi trên Moodle.",
            },
        }

    p1 = sv_data.get("p1") or {}
    p2 = sv_data.get("p2") or {}
    e1 = (p1.get("essay") or "").strip()
    e2 = (p2.get("essay") or "").strip()
    w1 = p1.get("word_count") if p1.get("word_count") is not None else count_words(e1)
    w2 = p2.get("word_count") if p2.get("word_count") is not None else count_words(e2)
    l1 = p1.get("line_count") or (len(e1.splitlines()) if e1 else 0)
    l2 = p2.get("line_count") or (len(e2.splitlines()) if e2 else 0)

    if kq is not None:
        kq = ensure_cogv_reasons(sv_data, kq)
        note = _detect_student_warning_tag(sv_data, kq)
        graded = True
    else:
        kq = KetQuaCham(
            p1_content=0, p1_org=0, p1_grammar=0, p1_vocab=0,
            p1_c_reason="Chưa chấm AI.", p1_o_reason="Chưa chấm AI.", p1_g_reason="Chưa chấm AI.", p1_v_reason="Chưa chấm AI.",
            p2_content=0, p2_org=0, p2_grammar=0, p2_vocab=0,
            p2_c_reason="Chưa chấm AI.", p2_o_reason="Chưa chấm AI.", p2_g_reason="Chưa chấm AI.", p2_v_reason="Chưa chấm AI.",
        )
        note = "Chưa chấm"
        graded = False

    return {
        "row": row_idx,
        "mssv": mssv,
        "ten": ten,
        "checked": bool(checked),
        "absent": False,
        "graded": graded,
        "note": note,
        "p1": {
            "code": p1.get("code") or "N/A",
            "prompt": clean_prompt_text(p1.get("prompt") or ""),
            "images": p1.get("images") or [],
            "essay": e1,
            "essay_lines": p1.get("essay_lines") or ([ln.strip() for ln in e1.splitlines() if ln.strip()] if e1 else []),
            "word_count": w1,
            "line_count": l1,
            "format_type": p1.get("format_type") or "BINH_THUONG",
            "c": int(kq.p1_content),
            "o": int(kq.p1_org),
            "g": int(kq.p1_grammar),
            "v": int(kq.p1_vocab),
            "rc": kq.p1_c_reason,
            "ro": kq.p1_o_reason,
            "rg": kq.p1_g_reason,
            "rv": kq.p1_v_reason,
        },
        "p2": {
            "code": p2.get("code") or "N/A",
            "prompt": clean_prompt_text(p2.get("prompt") or ""),
            "images": p2.get("images") or [],
            "essay": e2,
            "essay_lines": p2.get("essay_lines") or ([ln.strip() for ln in e2.splitlines() if ln.strip()] if e2 else []),
            "word_count": w2,
            "line_count": l2,
            "format_type": p2.get("format_type") or "BINH_THUONG",
            "c": int(kq.p2_content),
            "o": int(kq.p2_org),
            "g": int(kq.p2_grammar),
            "v": int(kq.p2_vocab),
            "rc": kq.p2_c_reason,
            "ro": kq.p2_o_reason,
            "rg": kq.p2_g_reason,
            "rv": kq.p2_v_reason,
        },
    }


def _refresh_saved_item_reasons(saved_it: dict, sv_data: dict | None = None) -> dict:
    """Đồng bộ lại đề bài và lý do chi tiết C-O-G-V của 1 sinh viên trong phiên _review.json."""
    if sv_data is not None:
        saved_it["p1"]["prompt"] = clean_prompt_text((sv_data.get("p1") or {}).get("prompt") or "")
        saved_it["p2"]["prompt"] = clean_prompt_text((sv_data.get("p2") or {}).get("prompt") or "")
        saved_it["p1"]["images"] = (sv_data.get("p1") or {}).get("images") or []
        saved_it["p2"]["images"] = (sv_data.get("p2") or {}).get("images") or []
    for pk, is_p1 in (("p1", True), ("p2", False)):
        p_obj = saved_it.get(pk) or {}
        p_obj["prompt"] = clean_prompt_text(p_obj.get("prompt") or "")
        rc, ro, rg, rv = _synthesize_part_cogv_reasons(
            p_obj["prompt"],
            p_obj.get("essay") or "",
            int(p_obj.get("c", 0)),
            int(p_obj.get("o", 0)),
            int(p_obj.get("g", 0)),
            int(p_obj.get("v", 0)),
            p_obj.get("rc") or "",
            p_obj.get("ro") or "",
            p_obj.get("rg") or "",
            p_obj.get("rv") or "",
            is_part1=is_p1,
            format_type=p_obj.get("format_type") or "BINH_THUONG",
        )
        p_obj["rc"], p_obj["ro"], p_obj["rg"], p_obj["rv"] = rc, ro, rg, rv
    return saved_it


def build_review_items_from_excel_and_json(
    excel_path: str = "",
    json_path: str = "",
    class_code: str = "",
    prefer_excel_scores: bool = False,
) -> list[dict]:
    """
    Khớp nối danh sách sinh viên từ file Excel, file JSON hoặc bộ nhớ tạm _review.json của lớp:
    - Chặn tuyệt đối việc ghép nhầm file JSON của lớp A với file Excel của lớp B.
    - Tự động khôi phục file cặp đôi nếu chỉ truyền 1 file (chỉ Excel hoặc chỉ JSON) để giảng viên luôn xem/sửa được cả Part 1 & Part 2.
    - Nếu file Excel vừa được giảng viên sửa trực tiếp (để khớp điểm với GV2) và lưu mới hơn file _review.json (hoặc prefer_excel_scores=True),
      tự động ưu tiên cập nhật điểm 8 cột C-O-G-V mới nhất từ file Excel.
    """
    validate_class_files_not_mixed(json_path, excel_path)
    json_p, excel_p = resolve_or_recover_class_files(class_code, json_path, excel_path)
    validate_class_files_not_mixed(json_p, excel_p)

    map_bai_lam: dict[str, dict] = {}
    ds_bai_lam: list[dict] = []
    if json_p and os.path.exists(json_p):
        with open(json_p, "r", encoding="utf-8") as f:
            raw_json = json.load(f)
        if isinstance(raw_json, dict) and isinstance(raw_json.get("items"), list):
            ds_bai_lam = raw_json["items"]
        elif isinstance(raw_json, list):
            ds_bai_lam = raw_json
        map_bai_lam = {
            normalize_mssv(item.get("mssv")): item
            for item in ds_bai_lam
            if normalize_mssv(item.get("mssv"))
        }

    code_hint = extract_explicit_iuh_class_code(json_p) or extract_explicit_iuh_class_code(excel_p)
    saved_session = load_review_session(json_p, excel_p, class_code=code_hint)
    saved_map: dict[str, dict] = {}
    if saved_session and isinstance(saved_session.get("items"), list):
        for it in saved_session["items"]:
            m = normalize_mssv(it.get("mssv"))
            if m:
                saved_map[m] = it

    # Trường hợp chỉ mới có file JSON (hoặc bộ nhớ _review.json) mà chưa có file Excel
    if not excel_p or not os.path.exists(excel_p):
        if saved_map:
            return [_refresh_saved_item_reasons(it, map_bai_lam.get(normalize_mssv(it.get("mssv")))) for it in saved_session["items"]]
        if ds_bai_lam:
            items_only_json = []
            for idx_j, item in enumerate(ds_bai_lam, start=10):
                m = normalize_mssv(item.get("mssv"))
                if not m:
                    continue
                ten = str(item.get("ho_ten") or item.get("ten") or m).strip()
                items_only_json.append(_make_review_item(idx_j, m, ten, item, None, checked=False))
            return items_only_json
        raise ValueError("Chưa tìm thấy dữ liệu bài làm hoặc bảng điểm của lớp này.")

    rev_path = get_review_session_path(json_p, excel_p)
    excel_is_newer = bool(prefer_excel_scores)
    if not excel_is_newer and excel_p and os.path.exists(excel_p):
        if not rev_path or not os.path.exists(rev_path):
            excel_is_newer = True
        else:
            try:
                if os.path.getmtime(excel_p) > os.path.getmtime(rev_path) + 0.1:
                    excel_is_newer = True
            except OSError:
                pass

    wb = openpyxl.load_workbook(excel_p, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    cols = detect_excel_schema(all_rows)
    c_mssv = cols["mssv"]
    c_ten = cols["ten"]
    c_ho_dem = cols["ho_dem"]
    p1_cols = cols["p1_cols"]
    p2_cols = cols["p2_cols"]
    c_comment = cols["comment"]

    if c_mssv is None:
        raise ValueError("Không nhận diện được cột MSSV trong file Excel.")

    # Kiểm tra trùng khớp MSSV giữa JSON và Excel nếu cả hai cùng có dữ liệu để tránh lẫn lộn 2 lớp khác nhau
    if map_bai_lam:
        excel_mssvs = {
            normalize_mssv(r[c_mssv])
            for r in all_rows
            if r and c_mssv < len(r) and normalize_mssv(r[c_mssv])
        }
        if excel_mssvs and len(excel_mssvs & set(map_bai_lam.keys())) == 0:
            raise ValueError(
                "CẢNH BÁO LẪN LỘN DỮ LIỆU GIỮA 2 LỚP:\n"
                "Không có bất kỳ Mã số sinh viên (MSSV) nào trùng khớp giữa File JSON và File Excel đã chọn!\n"
                "Hệ thống đã chặn ghép nối để bảo vệ dữ liệu gốc."
            )

    items: list[dict] = []
    for row_idx, r in enumerate(all_rows, 1):
        if not r or c_mssv >= len(r):
            continue
        mssv = normalize_mssv(r[c_mssv])
        if not mssv:
            continue
        if c_ho_dem is not None and c_ten is not None and c_ho_dem < len(r) and c_ten < len(r):
            ho = str(r[c_ho_dem] or "").strip()
            ten_r = str(r[c_ten] or "").strip()
            ten = f"{ho} {ten_r}".strip()
        elif c_ten is not None and c_ten < len(r) and r[c_ten] is not None:
            val_ten = str(r[c_ten]).strip()
            ten = val_ten if val_ten and not val_ten.startswith("=") else mssv
        else:
            ten = mssv

        sv_data = map_bai_lam.get(mssv)

        # Kiểm tra nếu trong file Excel đã có sẵn điểm C-O-G-V của sinh viên này
        kq_from_excel = None
        if p1_cols and p2_cols and max(p1_cols + p2_cols) < len(r):
            raw_scores = [r[idx] for idx in (p1_cols + p2_cols)]
            if all(v is not None and str(v).strip() != "" for v in raw_scores):
                try:
                    s_vals = [int(float(str(v).strip().replace(",", "."))) for v in raw_scores]
                    existing_cmt = ""
                    if c_comment is not None and c_comment < len(r) and r[c_comment] is not None:
                        existing_cmt = str(r[c_comment]).strip()
                    kq_from_excel = KetQuaCham(
                        p1_content=s_vals[0], p1_org=s_vals[1], p1_grammar=s_vals[2], p1_vocab=s_vals[3],
                        p2_content=s_vals[4], p2_org=s_vals[5], p2_grammar=s_vals[6], p2_vocab=s_vals[7],
                        nhan_xet=existing_cmt,
                    )
                    if sv_data is not None:
                        kq_from_excel = ensure_cogv_reasons(sv_data, kq_from_excel)
                except ValueError:
                    kq_from_excel = None

        if mssv in saved_map and (saved_map[mssv].get("graded") or sv_data is None):
            saved_it = saved_map[mssv]
            saved_it["row"] = row_idx
            saved_it["ten"] = ten
            if excel_is_newer and kq_from_excel is not None:
                p1_saved = saved_it.setdefault("p1", {})
                p2_saved = saved_it.setdefault("p2", {})
                old_p1 = (p1_saved.get("c"), p1_saved.get("o"), p1_saved.get("g"), p1_saved.get("v"))
                new_p1 = (kq_from_excel.p1_content, kq_from_excel.p1_org, kq_from_excel.p1_grammar, kq_from_excel.p1_vocab)
                if old_p1 != new_p1:
                    p1_saved["c"], p1_saved["o"], p1_saved["g"], p1_saved["v"] = new_p1
                    p1_saved["rc"] = p1_saved["ro"] = p1_saved["rg"] = p1_saved["rv"] = ""
                old_p2 = (p2_saved.get("c"), p2_saved.get("o"), p2_saved.get("g"), p2_saved.get("v"))
                new_p2 = (kq_from_excel.p2_content, kq_from_excel.p2_org, kq_from_excel.p2_grammar, kq_from_excel.p2_vocab)
                if old_p2 != new_p2:
                    p2_saved["c"], p2_saved["o"], p2_saved["g"], p2_saved["v"] = new_p2
                    p2_saved["rc"] = p2_saved["ro"] = p2_saved["rg"] = p2_saved["rv"] = ""
            items.append(_refresh_saved_item_reasons(saved_it, sv_data))
            continue

        if sv_data is None and kq_from_excel is not None:
            # Nếu chỉ mở mỗi file Excel đã có điểm mà chưa có file JSON, vẫn tạo bản ghi hiển thị được điểm 8 cột C-O-G-V
            dummy_sv = {
                "mssv": mssv,
                "ho_ten": ten,
                "p1": {"code": "ĐIỂM TỪ EXCEL", "prompt": "", "essay": "[Đang xem điểm trực tiếp từ file Excel]"},
                "p2": {"code": "ĐIỂM TỪ EXCEL", "prompt": "", "essay": "[Đang xem điểm trực tiếp từ file Excel]"},
            }
            items.append(_make_review_item(row_idx, mssv, ten, dummy_sv, kq_from_excel, checked=True))
        else:
            items.append(_make_review_item(row_idx, mssv, ten, sv_data, kq_from_excel, checked=False))

    return items


def chay_cham_ai_tao_ban_duyet(
    api_key: str,
    json_path: str,
    excel_path: str,
    rubric_text: str,
    progress_cb=None,
    stop_event: threading.Event | None = None,
    grader_fn=None,
    model_name: str = DEFAULT_MODEL,
    only_ungraded: bool = False,
    enable_pass2_arbitrator: bool = False,
    arbitrator_model: str = DEFAULT_ARBITRATOR_MODEL,
    stage_cb=None,
) -> dict:
    """Chấm AI theo lô và lưu kết quả vào phiên kiểm duyệt (_review.json), KHÔNG ghi đè trực tiếp vào Excel."""
    with open(json_path, "r", encoding="utf-8") as f:
        ds_bai_lam = json.load(f)
    map_bai_lam = {
        normalize_mssv(item.get("mssv")): item
        for item in ds_bai_lam
        if normalize_mssv(item.get("mssv"))
    }

    existing_items = build_review_items_from_excel_and_json(excel_path, json_path)
    items_by_mssv = {it["mssv"]: it for it in existing_items}

    tasks = []
    absent_rows = []
    for it in existing_items:
        mssv = it["mssv"]
        row_idx = it["row"]
        ten = it["ten"]
        if mssv not in map_bai_lam:
            absent_rows.append((row_idx, mssv, ten))
            continue
        if only_ungraded and it.get("graded"):
            continue
        tasks.append((row_idx, mssv, ten, map_bai_lam[mssv]))

    errors: list[str] = []
    graded_count = 0
    api_calls_used = 0

    if grader_fn is not None:
        with ThreadPoolExecutor(max_workers=8) as executor:
            future_to_info = {
                executor.submit(grader_fn, api_key, data, rubric_text): (row_idx, mssv, ten, data)
                for row_idx, mssv, ten, data in tasks
            }
            completed = 0
            for future in as_completed(future_to_info):
                if stop_event and stop_event.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                row_idx, mssv, ten, data = future_to_info[future]
                completed += 1
                try:
                    _, kq = future.result()
                    kq = ensure_cogv_reasons(data, kq)
                    items_by_mssv[mssv] = _make_review_item(row_idx, mssv, ten, data, kq, checked=False)
                    graded_count += 1
                    if progress_cb:
                        progress_cb(completed, len(tasks), mssv, ten, kq, None)
                except Exception as err:
                    errors.append(f"MSSV {mssv} ({ten}): {err}")
                    if progress_cb:
                        progress_cb(completed, len(tasks), mssv, ten, None, err)
    else:
        res_by_mssv, err_by_mssv, api_calls_used = _grade_tasks_by_prompt_clusters(
            api_key=api_key,
            tasks=tasks,
            rubric_text=rubric_text,
            model_name=model_name,
            stop_event=stop_event,
            enable_pass2_arbitrator=enable_pass2_arbitrator,
            arbitrator_model=arbitrator_model,
            stage_cb=stage_cb,
        )
        completed = 0
        for row_idx, mssv, ten, data in tasks:
            if stop_event and stop_event.is_set():
                break
            completed += 1
            kq = res_by_mssv.get(mssv)
            if kq is not None:
                kq = ensure_cogv_reasons(data, kq)
                rev_item = _make_review_item(row_idx, mssv, ten, data, kq, checked=False)
                c_meta = _LAST_COUNCIL_META_BY_MSSV.get(mssv) or {}
                if c_meta:
                    rev_item["council_stage"] = c_meta.get("stage", "")
                    rev_item["council_badge"] = c_meta.get("badge", "")
                    rev_item["council_verdict"] = c_meta.get("verdict", "")
                items_by_mssv[mssv] = rev_item
                graded_count += 1
                if progress_cb:
                    progress_cb(completed, len(tasks), mssv, ten, kq, None)
            else:
                err = err_by_mssv.get(mssv, RuntimeError("Không nhận được kết quả từ AI"))
                errors.append(f"MSSV {mssv} ({ten}): {err}")
                if progress_cb:
                    progress_cb(completed, len(tasks), mssv, ten, None, err)

    # Giữ nguyên tuyệt đối thứ tự dòng Excel (row) ban đầu của sinh viên trên giao diện
    final_items = [items_by_mssv[it["mssv"]] for it in existing_items]
    review_path = save_review_session(json_path, excel_path, final_items)
    return {
        "total_matched": len(tasks),
        "graded_count": graded_count,
        "absent_rows": absent_rows,
        "errors": errors,
        "review_path": review_path,
        "items": final_items,
        "api_calls_used": api_calls_used,
    }


def export_approved_scores_to_excel(
    excel_path: str,
    review_items: list[dict],
    include_comments: bool = False,
    allow_fallback: bool = True,
) -> dict:
    """Ghi 8 cột điểm C-O-G-V đã được giảng viên Approve sang file Excel Phiếu chấm điểm (không ghi nhận xét)."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    cols = detect_excel_schema(all_rows)

    c_mssv = cols["mssv"]
    p1_cols = cols["p1_cols"]
    p2_cols = cols["p2_cols"]
    c_comment = cols["comment"]

    if c_mssv is None or not p1_cols or not p2_cols:
        wb.close()
        raise ValueError("Không nhận diện được cấu trúc cột MSSV / Part 1 / Part 2 trong file Excel.")

    map_review = {
        normalize_mssv(it.get("mssv")): it
        for it in review_items
        if normalize_mssv(it.get("mssv"))
    }

    # Nếu không ghi nhận xét (mặc định khi Approve), xóa tiêu đề cột nhận xét AI nếu từng có
    if not include_comments and c_comment is not None:
        hdr_val = str(ws.cell(row=9, column=c_comment + 1).value or "").strip().lower()
        if "nhận xét" in hdr_val or "cảnh báo" in hdr_val or "ai comment" in hdr_val:
            ws.cell(row=9, column=c_comment + 1).value = None

    exported_count = 0
    for row_idx in range(1, ws.max_row + 1):
        cell_mssv = ws.cell(row=row_idx, column=c_mssv + 1).value
        mssv = normalize_mssv(cell_mssv)
        if not mssv or mssv not in map_review:
            continue
        it = map_review[mssv]
        if it.get("absent"):
            if not include_comments and c_comment is not None:
                ws.cell(row=row_idx, column=c_comment + 1).value = None
            continue

        p1 = it.get("p1") or {}
        p2 = it.get("p2") or {}
        ws.cell(row=row_idx, column=p1_cols[0] + 1, value=int(p1.get("c", 0)))
        ws.cell(row=row_idx, column=p1_cols[1] + 1, value=int(p1.get("o", 0)))
        ws.cell(row=row_idx, column=p1_cols[2] + 1, value=int(p1.get("g", 0)))
        ws.cell(row=row_idx, column=p1_cols[3] + 1, value=int(p1.get("v", 0)))

        ws.cell(row=row_idx, column=p2_cols[0] + 1, value=int(p2.get("c", 0)))
        ws.cell(row=row_idx, column=p2_cols[1] + 1, value=int(p2.get("o", 0)))
        ws.cell(row=row_idx, column=p2_cols[2] + 1, value=int(p2.get("g", 0)))
        ws.cell(row=row_idx, column=p2_cols[3] + 1, value=int(p2.get("v", 0)))

        if not include_comments and c_comment is not None:
            ws.cell(row=row_idx, column=c_comment + 1).value = None

        exported_count += 1

    if exported_count == 0 and review_items:
        wb.close()
        raise ValueError(
            "Không có sinh viên nào trong danh sách điểm hiện tại khớp MSSV với file Excel đích!\n"
            "Hệ thống đã tự động chặn ghi file để tránh ghi nhầm điểm sang lớp khác."
        )

    saved_path = excel_path
    try:
        wb.save(excel_path)
    except PermissionError:
        if not allow_fallback:
            wb.close()
            raise
        base, ext = os.path.splitext(excel_path)
        saved_path = f"{base}_APPROVED{ext}"
        wb.save(saved_path)
    wb.close()

    return {
        "exported_count": exported_count,
        "saved_path": saved_path,
        "cols": cols,
    }


def inspect_json_dataset(json_path: str) -> dict:
    """Đọc và thống kê nhanh file bai_lam_iuh.json để giảng viên kiểm tra Đề bài & Bài làm."""
    with open(json_path, "r", encoding="utf-8") as f:
        ds = json.load(f)
    if not isinstance(ds, list):
        raise ValueError("Cấu trúc file JSON không hợp lệ (phải là danh sách sinh viên).")

    p1_codes = set()
    p2_codes = set()
    p1_with_prompt = 0
    p2_with_prompt = 0
    p1_img_count = 0
    p2_img_count = 0
    non_empty_students = 0
    sample_p1_prompt = ""
    sample_p2_prompt = ""

    for item in ds:
        p1 = item.get("p1") or {}
        p2 = item.get("p2") or {}
        c1 = p1.get("code")
        c2 = p2.get("code")
        if c1 and c1 != "DA_CHAM_HOAC_TRONG":
            p1_codes.add(c1)
        if c2 and c2 != "DA_CHAM_HOAC_TRONG":
            p2_codes.add(c2)

        pr1 = (p1.get("prompt") or "").strip()
        pr2 = (p2.get("prompt") or "").strip()
        imgs1 = p1.get("images") or []
        imgs2 = p2.get("images") or []

        if pr1 or imgs1:
            p1_with_prompt += 1
            if pr1 and not sample_p1_prompt:
                sample_p1_prompt = pr1[:160].replace("\n", " ")
        if pr2 or imgs2:
            p2_with_prompt += 1
            if pr2 and not sample_p2_prompt:
                sample_p2_prompt = pr2[:160].replace("\n", " ")

        p1_img_count += len(imgs1)
        p2_img_count += len(imgs2)

        if count_words(p1.get("essay", "")) > 0 or count_words(p2.get("essay", "")) > 0:
            non_empty_students += 1

    return {
        "total": len(ds),
        "non_empty_students": non_empty_students,
        "p1_codes": sorted(p1_codes),
        "p2_codes": sorted(p2_codes),
        "p1_with_prompt": p1_with_prompt,
        "p2_with_prompt": p2_with_prompt,
        "p1_img_count": p1_img_count,
        "p2_img_count": p2_img_count,
        "sample_p1_prompt": sample_p1_prompt,
        "sample_p2_prompt": sample_p2_prompt,
    }


def chay_cham_8_luong_ghi_excel(
    api_key: str,
    json_path: str,
    excel_path: str,
    rubric_text: str,
    progress_cb=None,
    stop_event: threading.Event | None = None,
    grader_fn=None,
    model_name: str = DEFAULT_MODEL,
) -> dict:
    with open(json_path, "r", encoding="utf-8") as f:
        ds_bai_lam = json.load(f)
    map_bai_lam = {
        normalize_mssv(item.get("mssv")): item
        for item in ds_bai_lam
        if normalize_mssv(item.get("mssv"))
    }

    # Mở file Excel KHÔNG dùng data_only=True để giữ nguyên toàn bộ công thức tính tổng
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    cols = detect_excel_schema(all_rows)

    c_mssv = cols["mssv"]
    c_ten = cols["ten"]
    c_ho_dem = cols["ho_dem"]
    p1_cols = cols["p1_cols"]
    p2_cols = cols["p2_cols"]
    c_comment = cols["comment"]

    if c_mssv is None or not p1_cols or not p2_cols:
        wb.close()
        raise ValueError("Không nhận diện được cấu trúc cột MSSV / Part 1 / Part 2 trong file Excel.")

    header_comment = ws.cell(row=9, column=c_comment + 1).value
    if not header_comment:
        ws.cell(row=9, column=c_comment + 1, value="Nhận xét AI & Cảnh báo lạc đề")

    tasks = []
    absent_rows = []
    for row_idx in range(1, ws.max_row + 1):
        cell_mssv = ws.cell(row=row_idx, column=c_mssv + 1).value
        mssv = normalize_mssv(cell_mssv)
        if not mssv:
            continue
        if c_ho_dem is not None and c_ten is not None:
            ho = str(ws.cell(row=row_idx, column=c_ho_dem + 1).value or "").strip()
            ten_r = str(ws.cell(row=row_idx, column=c_ten + 1).value or "").strip()
            ten = f"{ho} {ten_r}".strip()
        elif c_ten is not None:
            val_ten = ws.cell(row=row_idx, column=c_ten + 1).value
            ten = str(val_ten).strip() if val_ten and not str(val_ten).startswith("=") else mssv
        else:
            ten = mssv

        if mssv in map_bai_lam:
            tasks.append((row_idx, mssv, ten, map_bai_lam[mssv]))
        else:
            absent_rows.append((row_idx, mssv, ten))

    results_map: dict[int, KetQuaCham] = {}
    errors: list[str] = []

    # Nếu truyền custom grader_fn (ví dụ trong Unit Test) thì chấm theo grader_fn
    if grader_fn is not None:
        with ThreadPoolExecutor(max_workers=8) as executor:
            future_to_info = {
                executor.submit(grader_fn, api_key, data, rubric_text): (row_idx, mssv, ten)
                for row_idx, mssv, ten, data in tasks
            }
            completed = 0
            for future in as_completed(future_to_info):
                if stop_event and stop_event.is_set():
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                row_idx, mssv, ten = future_to_info[future]
                completed += 1
                try:
                    _, kq = future.result()
                    results_map[row_idx] = kq
                    if progress_cb:
                        progress_cb(completed, len(tasks), mssv, ten, kq, None)
                except Exception as err:
                    errors.append(f"MSSV {mssv} ({ten}): {err}")
                    if progress_cb:
                        progress_cb(completed, len(tasks), mssv, ten, None, err)
    else:
        res_by_mssv, err_by_mssv, _ = _grade_tasks_by_prompt_clusters(
            api_key=api_key,
            tasks=tasks,
            rubric_text=rubric_text,
            model_name=model_name,
            stop_event=stop_event,
        )
        completed = 0
        for row_idx, mssv, ten, _ in tasks:
            if stop_event and stop_event.is_set():
                break
            completed += 1
            kq = res_by_mssv.get(mssv)
            if kq is not None:
                results_map[row_idx] = kq
                if progress_cb:
                    progress_cb(completed, len(tasks), mssv, ten, kq, None)
            else:
                err = err_by_mssv.get(mssv, RuntimeError("Không nhận được kết quả từ AI"))
                errors.append(f"MSSV {mssv} ({ten}): {err}")
                if progress_cb:
                    progress_cb(completed, len(tasks), mssv, ten, None, err)

    # Ghi điểm thành phần vào Excel (1-based column index cho openpyxl)
    for row_idx, kq in results_map.items():
        ws.cell(row=row_idx, column=p1_cols[0] + 1, value=kq.p1_content)
        ws.cell(row=row_idx, column=p1_cols[1] + 1, value=kq.p1_org)
        ws.cell(row=row_idx, column=p1_cols[2] + 1, value=kq.p1_grammar)
        ws.cell(row=row_idx, column=p1_cols[3] + 1, value=kq.p1_vocab)

        ws.cell(row=row_idx, column=p2_cols[0] + 1, value=kq.p2_content)
        ws.cell(row=row_idx, column=p2_cols[1] + 1, value=kq.p2_org)
        ws.cell(row=row_idx, column=p2_cols[2] + 1, value=kq.p2_grammar)
        ws.cell(row=row_idx, column=p2_cols[3] + 1, value=kq.p2_vocab)

        ws.cell(row=row_idx, column=c_comment + 1, value=kq.nhan_xet)

    for row_idx, mssv, ten in absent_rows:
        ws.cell(row=row_idx, column=c_comment + 1, value="⚠️ VẮNG THI (Không có bài trên hệ thống Moodle)")

    saved_path = excel_path
    try:
        wb.save(excel_path)
    except PermissionError:
        base, ext = os.path.splitext(excel_path)
        saved_path = f"{base}_AI_CHAM{ext}"
        wb.save(saved_path)
    wb.close()

    return {
        "total_matched": len(tasks),
        "graded_count": len(results_map),
        "absent_rows": absent_rows,
        "errors": errors,
        "cols": cols,
        "saved_path": saved_path,
    }


def extract_class_code(json_path: str = "", excel_path: str = "", fallback_name: str = "") -> str:
    """Tự động trích xuất Mã lớp (VD: 2607AVCB005, 2606AVCB034) từ tên file JSON/Excel hoặc tên thư mục."""
    candidates = []
    for p in (json_path, excel_path):
        if p:
            candidates.append(os.path.basename(p))
            parent = os.path.basename(os.path.dirname(os.path.abspath(p)))
            if parent:
                candidates.append(parent)
    if fallback_name:
        candidates.append(fallback_name)

    for text in candidates:
        m = re.search(r"(\d{4}\s*[A-Z]{3,8}\.?\s*\d{2,4})", text, flags=re.IGNORECASE)
        if m:
            return re.sub(r"[\s.]+", "", m.group(1)).upper()

    for text in candidates:
        base = os.path.splitext(text)[0]
        cleaned = re.sub(
            r"(?i)^(bai_lam_iuh_|bai_lam_iuh|phiếu\s*chấm\s*viết\s*cuối\s*kỳ\s*|phieu\s*cham\s*viet\s*cuoi\s*ky\s*)",
            "",
            base,
        )
        cleaned = re.sub(r"(?i)(_review|_updated)$", "", cleaned).strip(" _-.")
        if cleaned and cleaned.lower() not in ("src", "auto_nhap_diem_iuh", "desktop", "downloads", "documents"):
            return cleaned[:22].upper()

    return "LOP_MOI"


def auto_pair_class_files(selected_path: str) -> tuple[str, str]:
    """Tự động tìm và ghép cặp file JSON <-> Excel cùng lớp trong cùng thư mục khi mở thêm lớp mới."""
    if not selected_path:
        return "", ""
    abs_path = os.path.abspath(selected_path)
    folder = os.path.dirname(abs_path)
    ext = os.path.splitext(abs_path)[1].lower()
    code = extract_class_code(
        json_path=abs_path if ext == ".json" else "",
        excel_path=abs_path if ext == ".xlsx" else "",
    )

    if ext == ".json":
        json_path = abs_path
        if json_path.lower().endswith("_review.json"):
            raw_candidate = json_path[:-12] + ".json"
            if os.path.exists(raw_candidate):
                json_path = raw_candidate
        excel_path = ""
        if os.path.isdir(folder):
            xlsx_files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(".xlsx") and not f.startswith("~$")
            ]
            matched = [p for p in xlsx_files if code and code.lower() in os.path.basename(p).lower()]
            pool = matched or xlsx_files
            if pool:
                non_updated = [p for p in pool if not os.path.basename(p).lower().endswith("_updated.xlsx")]
                target_pool = non_updated or pool
                target_pool.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
                excel_path = target_pool[0]
        return json_path, excel_path

    if ext == ".xlsx":
        excel_path = abs_path
        json_path = ""
        if os.path.isdir(folder):
            json_files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(".json")
                and not f.lower().endswith("_review.json")
                and f.lower() != "config.json"
            ]
            matched = [p for p in json_files if code and code.lower() in os.path.basename(p).lower()]
            pool = matched or json_files
            if pool:
                pool.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
                json_path = pool[0]
        return json_path, excel_path

    return "", ""


def build_js_paced_grade_submitter(
    grades_map: dict[str, dict],
    min_delay_sec: float = 3.0,
    max_delay_sec: float = 5.0,
    class_codes: list[str] | None = None,
) -> str:
    """Tạo đoạn mã JS F12 chạy ngầm tuần tự (1 luồng + nghỉ tự nhiên 3-5s/câu, tự động bỏ qua câu đã được người khác chấm)."""
    min_ms = max(500, int(round(min_delay_sec * 1000)))
    max_ms = max(min_ms, int(round(max_delay_sec * 1000)))
    grades_json = json.dumps(grades_map, ensure_ascii=False)
    classes_label = ", ".join(class_codes) if class_codes else "Lớp hiện tại"

    return f"""(async () => {{
    window.__iuhStopGrading = false;
    const GRADES_MAP = {grades_json};
    const MIN_DELAY_MS = {min_ms};
    const MAX_DELAY_MS = {max_ms};
    const CLASSES_LABEL = {json.dumps(classes_label, ensure_ascii=False)};

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const randomDelay = () => MIN_DELAY_MS + Math.floor(Math.random() * (MAX_DELAY_MS - MIN_DELAY_MS + 1));

    function isSlotAlreadyGradedOnTable(linkEl) {{
        if (!linkEl) return false;
        const txt = (linkEl.textContent || "").trim();
        if (!txt) return false;
        if (/yêu\\s*cầu\\s*chấm\\s*điểm|requires\\s*grading|not\\s*yet\\s*graded|chưa\\s*chấm/i.test(txt)) {{
            return false;
        }}
        return /\\d/.test(txt);
    }}

    // Quét danh sách sinh viên trên trang Bảng điểm Moodle hiện tại (khớp chính xác theo MSSV từng dòng, không bao giờ lệch dòng)
    const rows = Array.from(document.querySelectorAll('table tbody tr'));
    const rowMap = new Map();
    for (const tr of rows) {{
        if (tr.classList.contains('emptyrow')) continue;
        const text = tr.innerText || "";
        const svMatch = text.match(/sv_(\\d{{6,12}})/i);
        const emailMatch = text.match(/\\b(\\d{{6,12}})@[a-z0-9._-]+/i);
        const links = Array.from(tr.querySelectorAll('a[href*="reviewquestion.php"], a[href*="comment.php"]'));

        let mssv = null;
        if (svMatch) {{
            mssv = svMatch[1];
        }} else if (emailMatch) {{
            mssv = emailMatch[1];
        }} else {{
            const cells = Array.from(tr.querySelectorAll('td, th')).map(td => (td.innerText || "").trim());
            const exactCell = cells.find(c => /^[12]\\d{{7}}$/.test(c));
            if (exactCell) {{
                mssv = exactCell;
            }} else if (links.length > 0) {{
                const fallbackMatch = text.match(/\\b([12]\\d{{7}})\\b/);
                if (fallbackMatch) mssv = fallbackMatch[1];
            }}
        }}
        if (!mssv || !GRADES_MAP[mssv]) continue;

        const q17El = links.find(a => a.href.includes('slot=17')) || null;
        const q18El = links.find(a => a.href.includes('slot=18')) || null;
        if (!q17El && !q18El) continue;

        const prev = rowMap.get(mssv);
        rowMap.set(mssv, {{
            mssv,
            ten: GRADES_MAP[mssv].ten || "",
            class_code: GRADES_MAP[mssv].class_code || "",
            q17Score: GRADES_MAP[mssv].q17,
            q18Score: GRADES_MAP[mssv].q18,
            q17El: q17El || (prev ? prev.q17El : null),
            q18El: q18El || (prev ? prev.q18El : null),
            tr
        }});
    }}

    const matchedStudents = Array.from(rowMap.values());
    const questionTasks = [];
    let alreadyGradedOnTableCount = 0;

    for (const st of matchedStudents) {{
        if (st.q17El && st.q17Score !== undefined && st.q17Score !== null && String(st.q17Score).trim() !== "") {{
            const gradedTbl = isSlotAlreadyGradedOnTable(st.q17El);
            if (gradedTbl) alreadyGradedOnTableCount++;
            questionTasks.push({{
                mssv: st.mssv,
                ten: st.ten,
                slotLabel: "Q.17 (Part 1)",
                score: String(st.q17Score).trim(),
                linkEl: st.q17El,
                url: st.q17El.href,
                alreadyGradedOnTable: gradedTbl,
                tableGradeText: (st.q17El.textContent || "").trim()
            }});
        }}
        if (st.q18El && st.q18Score !== undefined && st.q18Score !== null && String(st.q18Score).trim() !== "") {{
            const gradedTbl = isSlotAlreadyGradedOnTable(st.q18El);
            if (gradedTbl) alreadyGradedOnTableCount++;
            questionTasks.push({{
                mssv: st.mssv,
                ten: st.ten,
                slotLabel: "Q.18 (Part 2)",
                score: String(st.q18Score).trim(),
                linkEl: st.q18El,
                url: st.q18El.href,
                alreadyGradedOnTable: gradedTbl,
                tableGradeText: (st.q18El.textContent || "").trim()
            }});
        }}
    }}

    if (questionTasks.length === 0) {{
        alert("Không tìm thấy sinh viên nào trên trang Moodle này khớp với danh sách điểm đã nạp (" + CLASSES_LABEL + ").\\nHãy kiểm tra xem bạn đã mở đúng trang Bảng điểm của lớp chưa.");
        return;
    }}

    // Khởi tạo bảng trạng thái nổi phong cách Apple trên góc phải màn hình Moodle
    const oldBanner = document.getElementById('iuh-auto-grade-banner');
    if (oldBanner) oldBanner.remove();

    const banner = document.createElement('div');
    banner.id = 'iuh-auto-grade-banner';
    banner.style.cssText = 'position:fixed;top:16px;right:16px;z-index:999999;width:430px;background:#ffffff;color:#1d1d1f;border:1px solid #d2d2d7;border-radius:12px;padding:14px 16px;box-shadow:0 16px 36px rgba(0,0,0,0.14);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:13px;';
    banner.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
            <strong style="color:#1d1d1f;font-size:13.5px;letter-spacing:-0.01em;">Nhập điểm Moodle chạy ngầm (3–5s/câu)</strong>
            <div style="display:flex;gap:6px;">
                <button id="iuh-btn-stop-grade" style="background:#f5f5f7;color:#d70015;border:1px solid #d2d2d7;border-radius:6px;padding:3px 10px;cursor:pointer;font-weight:600;font-size:12px;">Dừng</button>
                <button id="iuh-btn-close-grade" style="background:#f5f5f7;color:#6e6e73;border:1px solid #d2d2d7;border-radius:6px;padding:3px 8px;cursor:pointer;font-size:12px;">Đóng</button>
            </div>
        </div>
        <div id="iuh-grade-sub" style="color:#6e6e73;font-size:12px;margin-bottom:8px;">
            Khớp <strong style="color:#1d1d1f;">${{matchedStudents.length}} sinh viên</strong> (${{questionTasks.length}} câu hỏi · Đã có điểm sẵn: ${{alreadyGradedOnTableCount}} câu).
        </div>
        <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:#1d1d1f;margin-bottom:8px;cursor:pointer;user-select:none;">
            <input type="checkbox" id="iuh-chk-overwrite-existing" style="cursor:pointer;" />
            <span>Ghi đè cả những câu đã có người chấm trước đó</span>
        </label>
        <div style="width:100%;height:6px;background:#e5e5ea;border-radius:3px;overflow:hidden;margin-bottom:8px;">
            <div id="iuh-grade-bar" style="width:0%;height:100%;background:#0071e3;transition:width 0.25s;"></div>
        </div>
        <div id="iuh-grade-status" style="color:#1d1d1f;font-weight:500;font-size:12.5px;">Chuẩn bị bắt đầu...</div>
    `;
    document.body.appendChild(banner);

    const statusEl = document.getElementById('iuh-grade-status');
    const barEl = document.getElementById('iuh-grade-bar');
    const stopBtn = document.getElementById('iuh-btn-stop-grade');
    const closeBtn = document.getElementById('iuh-btn-close-grade');
    const overwriteChk = document.getElementById('iuh-chk-overwrite-existing');

    stopBtn.onclick = () => {{
        window.__iuhStopGrading = true;
        stopBtn.disabled = true;
        stopBtn.textContent = "Đang dừng...";
        statusEl.textContent = "Đang dừng sau câu hiện tại...";
    }};
    closeBtn.onclick = () => {{
        window.__iuhStopGrading = true;
        banner.remove();
    }};

    async function postGradeForm(form, commentUrl, markInput, scoreStr) {{
        const actionAttr = form.getAttribute('action') || commentUrl;
        const postUrl = new URL(actionAttr, commentUrl).href;
        const params = new URLSearchParams();

        const elements = Array.from(form.querySelectorAll('input, textarea, select'));
        let hasSubmitParam = false;
        for (const el of elements) {{
            const name = el.getAttribute('name');
            if (!name || el.disabled) continue;
            const type = (el.getAttribute('type') || '').toLowerCase();
            if (type === 'cancel' || type === 'button' || type === 'reset' || type === 'file' || type === 'image') continue;
            if ((type === 'checkbox' || type === 'radio') && !el.checked) continue;
            if (name === 'cancel') continue;

            if (el === markInput || name.endsWith('-mark')) {{
                params.set(name, String(scoreStr));
            }} else if (type === 'submit') {{
                if (name === 'submit' || !hasSubmitParam) {{
                    params.set(name, el.value || 'Lưu');
                    hasSubmitParam = true;
                }}
            }} else if (el.tagName.toLowerCase() === 'textarea') {{
                params.set(name, el.value || el.textContent || '');
            }} else {{
                params.set(name, el.value ?? '');
            }}
        }}
        if (!hasSubmitParam) {{
            params.set('submit', '1');
        }}

        const postResp = await fetch(postUrl, {{
            method: 'POST',
            credentials: 'include',
            headers: {{ 'Content-Type': 'application/x-www-form-urlencoded' }},
            body: params.toString()
        }});
        if (!postResp.ok) {{
            throw new Error("HTTP " + postResp.status + " khi lưu điểm");
        }}
        const postHtml = await postResp.text();
        const hasFormError = postHtml.includes('-mark') && /id_error_|invalid-feedback|felement\\s+error/i.test(postHtml);
        return !hasFormError;
    }}

    async function submitSingleSlot(task) {{
        const rawUrl = task.url;
        const scoreValue = task.score;
        const allowOverwrite = Boolean(overwriteChk && overwriteChk.checked);

        // Chốt bảo vệ 1: Nếu trên bảng điểm câu này đã có điểm sẵn và đang ở chế độ Mặc định An toàn (không ghi đè) -> Bỏ qua ngay (0s)!
        if (task.alreadyGradedOnTable && !allowOverwrite) {{
            return {{ status: 'skipped_existing', existingMark: task.tableGradeText }};
        }}

        let commentUrl = rawUrl.replace('reviewquestion.php', 'comment.php');
        let resp = await fetch(commentUrl, {{ credentials: 'include' }});
        if (!resp.ok) throw new Error("HTTP " + resp.status + " khi mở form chấm");
        let html = await resp.text();
        let doc = new DOMParser().parseFromString(html, 'text/html');

        let markInput = doc.querySelector('input[name$="-mark"], input[id$="-mark"]');
        if (!markInput) {{
            // Dự phòng nếu đường dẫn comment.php nằm bên trong trang reviewquestion.php
            const revResp = await fetch(rawUrl, {{ credentials: 'include' }});
            const revHtml = await revResp.text();
            const revDoc = new DOMParser().parseFromString(revHtml, 'text/html');
            const cLink = revDoc.querySelector('a[href*="comment.php"]');
            if (cLink) {{
                commentUrl = new URL(cLink.getAttribute('href'), rawUrl).href;
                resp = await fetch(commentUrl, {{ credentials: 'include' }});
                html = await resp.text();
                doc = new DOMParser().parseFromString(html, 'text/html');
                markInput = doc.querySelector('input[name$="-mark"], input[id$="-mark"]');
            }}
        }}
        if (!markInput) {{
            throw new Error("Không tìm thấy ô nhập điểm (-mark) trong form Moodle");
        }}

        // Chốt bảo vệ 2 (Trực tiếp từ Server Moodle): Kiểm tra xem giữa chừng có giảng viên khác vừa chấm câu này chưa!
        const existingMarkRaw = (markInput.getAttribute('value') ?? markInput.value ?? "").trim();
        const existingNum = parseFloat(existingMarkRaw.replace(',', '.'));
        const targetNum = parseFloat(String(scoreValue).replace(',', '.'));
        const hasExistingServerGrade = existingMarkRaw !== "" && !Number.isNaN(existingNum);

        if (hasExistingServerGrade) {{
            if (!allowOverwrite) {{
                return {{ status: 'skipped_existing', existingMark: existingMarkRaw }};
            }}
            if (!Number.isNaN(targetNum) && Math.abs(existingNum - targetNum) < 0.001) {{
                return {{ status: 'skipped_same', existingMark: existingMarkRaw }};
            }}
        }}

        const form = markInput.closest('form') || doc.querySelector('form[action*="comment.php"], form.mform, form');
        if (!form) {{
            throw new Error("Không tìm thấy biểu mẫu <form> chấm điểm");
        }}

        let ok = await postGradeForm(form, commentUrl, markInput, scoreValue);
        if (!ok && (String(scoreValue).includes('.') || String(scoreValue).includes(','))) {{
            const altScore = String(scoreValue).includes('.')
                ? String(scoreValue).replace('.', ',')
                : String(scoreValue).replace(',', '.');
            ok = await postGradeForm(form, commentUrl, markInput, altScore);
        }}
        if (!ok) {{
            throw new Error("Moodle từ chối lưu giá trị điểm " + scoreValue);
        }}
        return {{ status: 'saved', existingMark: existingMarkRaw }};
    }}

    let successCount = 0;
    let skippedCount = 0;
    let errorCount = 0;

    for (let i = 0; i < questionTasks.length; i++) {{
        if (window.__iuhStopGrading) break;
        const t = questionTasks[i];
        const pct = Math.round(((i) / questionTasks.length) * 100);
        barEl.style.width = pct + "%";
        statusEl.textContent = `[${{i + 1}}/${{questionTasks.length}}] Đang kiểm tra ${{t.mssv}} (${{t.ten}}) — ${{t.slotLabel}}...`;

        let didPostToServer = false;
        try {{
            const res = await submitSingleSlot(t);
            if (res.status === 'skipped_existing' || res.status === 'skipped_same') {{
                skippedCount++;
                if (t.linkEl) {{
                    t.linkEl.textContent = `Đã có: ${{res.existingMark}} (Giữ nguyên)`;
                    t.linkEl.style.background = '#f5f5f7';
                    t.linkEl.style.color = '#6e6e73';
                    t.linkEl.style.padding = '2px 6px';
                    t.linkEl.style.borderRadius = '4px';
                }}
                console.log(`[Bỏ qua ${{i + 1}}/${{questionTasks.length}}] ${{t.mssv}} (${{t.ten}}) - ${{t.slotLabel}} đã có điểm (${{res.existingMark}}), giữ nguyên không ghi đè.`);
            }} else {{
                didPostToServer = true;
                successCount++;
                if (t.linkEl) {{
                    t.linkEl.textContent = `Đã lưu: ${{t.score}}`;
                    t.linkEl.style.background = '#e8f5e9';
                    t.linkEl.style.color = '#1b5e20';
                    t.linkEl.style.fontWeight = '600';
                    t.linkEl.style.padding = '2px 6px';
                    t.linkEl.style.borderRadius = '4px';
                }}
                console.log(`[Đã lưu ${{i + 1}}/${{questionTasks.length}}] ${{t.mssv}} (${{t.ten}}) - ${{t.slotLabel}} = ${{t.score}}`);
            }}
        }} catch (err) {{
            errorCount++;
            if (t.linkEl) {{
                t.linkEl.style.background = '#ffebee';
                t.linkEl.style.color = '#b71c1c';
            }}
            console.error(`[Lỗi ${{i + 1}}/${{questionTasks.length}}] Khi nhập ${{t.mssv}} - ${{t.slotLabel}}:`, err);
        }}

        const nextPct = Math.round(((i + 1) / questionTasks.length) * 100);
        barEl.style.width = nextPct + "%";

        if (didPostToServer && i < questionTasks.length - 1 && !window.__iuhStopGrading) {{
            const waitMs = randomDelay();
            statusEl.textContent = `Đã lưu ${{t.mssv}} (${{t.slotLabel}} = ${{t.score}}). Nghỉ tự nhiên ${{(waitMs / 1000).toFixed(1)}}s...`;
            await sleep(waitMs);
        }}
    }}

    stopBtn.style.display = 'none';
    if (window.__iuhStopGrading) {{
        statusEl.style.color = '#b45309';
        statusEl.textContent = `Đã dừng tại câu ${{successCount + skippedCount + errorCount}}/${{questionTasks.length}} (Đã lưu mới: ${{successCount}}, Giữ nguyên: ${{skippedCount}}, Lỗi: ${{errorCount}}).`;
    }} else {{
        statusEl.style.color = '#1d1d1f';
        statusEl.innerHTML = `Hoàn tất: Đã lưu mới <b>${{successCount}}</b> câu · Giữ nguyên câu đã chấm: <b>${{skippedCount}}</b> · Lỗi: <b>${{errorCount}}</b>. <button id="iuh-btn-reload" style="margin-left:8px;background:#0071e3;color:#fff;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-weight:600;font-size:12px;">Tải lại trang</button>`;
        const reloadBtn = document.getElementById('iuh-btn-reload');
        if (reloadBtn) reloadBtn.onclick = () => location.reload();
    }}
}})();"""


class IUHWritingAISuiteApp:
    SCORE_SLOTS: list[tuple[str, str]] = [
        ("p1", "c"),
        ("p1", "o"),
        ("p1", "g"),
        ("p1", "v"),
        ("p2", "c"),
        ("p2", "o"),
        ("p2", "g"),
        ("p2", "v"),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("IUH Writing Grading Suite — Hệ Thống Chấm, Kiểm Duyệt & Nhập Điểm Writing")
        self.root.geometry("1280x820")
        self.root.minsize(820, 620)

        self.cfg = load_config()
        self.stop_ai_event = threading.Event()
        self.stop_web_event = threading.Event()
        self.is_running_ai = False
        self.is_running_web = False

        self.review_items: list[dict] = []
        self.filtered_indices: list[int] = []
        self.current_review_pos: int = 0
        self.focus_zone: str = "list"
        self.active_score_slot_idx: int = 0
        self.prompt_visible = {"p1": True, "p2": True}
        self._updating_score_vars = False
        self.rule_profile: str = resolve_rule_profile_key(self.cfg.get("rule_profile"))
        self.rubric_content = load_marking_rules(self.rule_profile)

        # Bộ nhớ phiên độc lập cho từng tab lớp học (Per-Class Session Store) & Hàng đợi chấm AI đa lớp
        self.class_sessions: dict[str, dict] = {}
        self.ai_queue: list[tuple[str, bool]] = []

        # Quản lý đa lớp học (Menu Lớp Dọc Nổi Floating Popover — 0px khoảng trống thừa)
        self.class_bar_visible: bool = bool(self.cfg.get("class_bar_visible", True))
        self.class_overlay_visible: bool = False
        self.pipeline_hud_visible: bool = False
        self.pipeline_step_widgets: dict[int, dict] = {}
        self.enable_pass2_var = tk.BooleanVar(value=bool(self.cfg.get("enable_pass2_arbitrator", True)))
        saved_arb_model = str(self.cfg.get("arbitrator_model") or DEFAULT_ARBITRATOR_MODEL).strip()
        if saved_arb_model not in AVAILABLE_MODELS:
            saved_arb_model = DEFAULT_ARBITRATOR_MODEL
        self.arbitrator_model_var = tk.StringVar(value=saved_arb_model)
        self.font_size_pt: int = max(9, min(16, int(self.cfg.get("font_size_pt", 11) or 11)))
        self.settings_win: tk.Toplevel | None = None
        self.open_classes: list[dict] = []
        raw_open = self.cfg.get("open_classes")
        if isinstance(raw_open, list) and raw_open:
            for entry in raw_open:
                if isinstance(entry, dict) and entry.get("code"):
                    c_code = str(entry.get("code")).strip()
                    j_p, e_p = resolve_or_recover_class_files(
                        c_code,
                        str(entry.get("json_path") or "").strip(),
                        str(entry.get("excel_path") or "").strip(),
                    )
                    self.open_classes.append({
                        "code": c_code,
                        "json_path": j_p,
                        "excel_path": e_p,
                        "checked_count": int(entry.get("checked_count") or 0),
                        "total_count": int(entry.get("total_count") or 0),
                    })
        init_json = str(self.cfg.get("last_json_path") or "").strip()
        init_excel = str(self.cfg.get("last_excel_path") or "").strip()
        if not self.open_classes and (init_json or init_excel):
            init_code = extract_class_code(init_json, init_excel)
            j_p, e_p = resolve_or_recover_class_files(init_code, init_json, init_excel)
            self.open_classes.append({
                "code": init_code,
                "json_path": j_p,
                "excel_path": e_p,
                "checked_count": 0,
                "total_count": 0,
            })
        saved_active = str(self.cfg.get("active_class_code") or "").strip()
        if saved_active and any(c["code"] == saved_active for c in self.open_classes):
            self.active_class_code = saved_active
        elif self.open_classes:
            self.active_class_code = self.open_classes[0]["code"]
        else:
            self.active_class_code = ""

        self._build_ui()
        self._bind_review_shortcuts()
        self._update_focus_visuals()
        self.root.after(150, self._initial_load)

    def _build_ui(self):
        self.root.configure(bg="#F5F5F7")
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")

        style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=(12, 4))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#1D1D1F")
        style.configure("Treeview.Heading", font=("Segoe UI", 8, "bold"), foreground="#6E6E73")

        # Khởi tạo các khung nền tảng (không pack topbar đen và không pack thanh 3 tab để đạt Zero-Header y=0 chuẩn Apple)
        topbar = tk.Frame(self.root, bg="#F5F5F7", height=0)
        self.topbar = topbar
        self.class_bar_frame = tk.Frame(topbar, bg="#F5F5F7")
        self.class_tabs_holder = tk.Frame(self.class_bar_frame, bg="#F5F5F7")
        self.class_tabs_holder.pack(side=tk.LEFT, fill=tk.Y)
        self.btn_add_class_tab = tk.Button(self.class_bar_frame, text="Mở lớp", command=self._add_class_tab_dialog)
        self.btn_add_class_tab.pack(side=tk.LEFT)
        self.btn_toggle_class_bar = tk.Button(topbar, text="Lớp học", command=self._toggle_class_overlay)
        self.btn_toggle_class_bar.pack(side=tk.LEFT)

        self.notebook = ttk.Notebook(self.root)
        self.tab1 = ttk.Frame(self.notebook, padding=8)
        self.tab3 = ttk.Frame(self.notebook, padding=8)

        # Màn hình chính duy nhất (Single-Screen Workspace) bắt đầu ngay tại đỉnh cửa sổ (y = 0)
        self.tab2 = tk.Frame(self.root, bg="#F5F5F7")
        self.tab2.pack(fill=tk.BOTH, expand=True)

        self._build_tab1()
        self._build_tab2()
        self._build_tab3()

        # Khung Menu Lớp Dọc Nổi (Apple Floating Popover — dùng .place() đè lên trên, 0px chiếm chỗ cố định)
        self.class_overlay_panel = tk.Frame(
            self.root,
            bg="#FFFFFF",
            highlightbackground="#D2D2D7",
            highlightcolor="#D2D2D7",
            highlightthickness=1,
            padx=10,
            pady=10,
        )
        self.root.bind("<Button-1>", self._on_root_click_close_overlay, add="+")

        self._apply_class_bar_visibility()
        self._render_class_bar()
        self._apply_font_size_to_widgets()

    # ==================== TAB 1: DỮ LIỆU & CHẤM AI ====================
    def _build_tab1(self):
        step1_frame = ttk.LabelFrame(
            self.tab1,
            text="Bước 1: Cấu hình Gemini API Key & Lấy dữ liệu bài làm từ trình duyệt (F12)",
            padding=10,
        )
        step1_frame.pack(fill=tk.X, pady=(0, 8))

        api_row = ttk.Frame(step1_frame)
        api_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(api_row, text="Gemini API Key:", width=16).pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar(value=self.cfg.get("api_key", ""))
        self.api_entry = ttk.Entry(api_row, textvariable=self.api_key_var, show="*", width=36)
        self.api_entry.pack(side=tk.LEFT, padx=(0, 6), fill=tk.X, expand=True)

        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            api_row,
            text="Hiện Key",
            variable=self.show_key_var,
            command=self._toggle_api_key_visibility,
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(api_row, text="Model AI:").pack(side=tk.LEFT, padx=(0, 4))
        saved_model = self.cfg.get("model_name", DEFAULT_MODEL)
        if saved_model not in AVAILABLE_MODELS or "lite" in saved_model or saved_model == "gemini-3.5-flash":
            saved_model = DEFAULT_MODEL
            save_config({"model_name": DEFAULT_MODEL})
        self.model_var = tk.StringVar(value=saved_model)
        self.model_combo = ttk.Combobox(
            api_row,
            textvariable=self.model_var,
            values=AVAILABLE_MODELS,
            state="readonly",
            width=22,
        )
        self.model_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(api_row, text="Lưu cấu hình", command=self._save_api_key).pack(side=tk.LEFT)

        scraper_row = ttk.Frame(step1_frame)
        scraper_row.pack(fill=tk.X)
        ttk.Button(
            scraper_row,
            text="Copy mã JS Hút Bài 8 Luồng (Dán vào Console F12 trên trình duyệt)",
            command=self._copy_js_scraper,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        step2_frame = ttk.LabelFrame(
            self.tab1,
            text="Bước 2: Chọn File Bài Làm JSON (F12) & File Excel Phiếu Chấm Điểm (.xlsx)",
            padding=10,
        )
        step2_frame.pack(fill=tk.X, pady=(0, 8))

        json_row = ttk.Frame(step2_frame)
        json_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(json_row, text="File Bài Làm (.json):", width=20).pack(side=tk.LEFT)
        self.json_path_var = tk.StringVar(value=self.cfg.get("last_json_path", ""))
        ttk.Entry(json_row, textvariable=self.json_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(json_row, text="Chọn File JSON...", command=self._browse_json).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(json_row, text="Kiểm tra Đề trong JSON", command=self._check_json_info).pack(side=tk.LEFT)

        excel1_row = ttk.Frame(step2_frame)
        excel1_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(excel1_row, text="File Excel Phiếu Chấm:", width=20).pack(side=tk.LEFT)
        self.excel1_path_var = tk.StringVar(value=self.cfg.get("last_excel_path", ""))
        ttk.Entry(excel1_row, textvariable=self.excel1_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(excel1_row, text="Chọn File Excel...", command=self._browse_excel_tab1).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(excel1_row, text="Kiểm tra Khớp nối", command=self._refresh_data_matching).pack(side=tk.LEFT)

        self.match_summary_var = tk.StringVar(
            value="Chưa nạp dữ liệu đối chiếu. Hãy chọn File JSON F12 và File Excel Phiếu chấm."
        )
        ttk.Label(
            step2_frame,
            textvariable=self.match_summary_var,
            font=("Segoe UI", 9, "bold"),
            foreground="#1e3a8a",
        ).pack(anchor=tk.W, pady=(2, 0))

        step3_frame = ttk.LabelFrame(
            self.tab1,
            text="Bước 3: Bộ Quy Tắc Chấm Theo Sách & Đợt Thi + Điều Khiển Chấm AI",
            padding=10,
        )
        step3_frame.pack(fill=tk.X, pady=(0, 8))

        rules_row = ttk.Frame(step3_frame)
        rules_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(rules_row, text="Bộ quy tắc chấm:").pack(side=tk.LEFT, padx=(0, 6))
        self.rule_profile_label_var = tk.StringVar(
            value=MARKING_RULE_PROFILES[self.rule_profile]["label"]
        )
        profile_labels = [p["label"] for p in MARKING_RULE_PROFILES.values()]
        self.rule_profile_combo = ttk.Combobox(
            rules_row,
            textvariable=self.rule_profile_label_var,
            values=profile_labels,
            state="readonly",
            width=26,
        )
        self.rule_profile_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.rule_profile_combo.bind("<<ComboboxSelected>>", self._on_rule_profile_selected)

        self.rules_info_var = tk.StringVar(value=self._format_rules_info_text())
        ttk.Label(rules_row, textvariable=self.rules_info_var, foreground="#334155").pack(side=tk.LEFT)
        ttk.Button(rules_row, text="Tải lại quy tắc", command=self._reset_rubric).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(rules_row, text="Mở file quy tắc", command=self._open_marking_rules_file).pack(side=tk.RIGHT)

        action_frame = ttk.Frame(step3_frame)
        action_frame.pack(fill=tk.X)

        self.btn_run_ai = tk.Button(
            action_frame,
            text="Chạy AI Chấm Toàn Bộ & Mở Bảng Kiểm Duyệt (Tab 2)",
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#1e40af",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
            command=lambda: self._start_ai_grading(only_ungraded=False),
        )
        self.btn_run_ai.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.btn_run_ungraded = ttk.Button(
            action_frame,
            text="Chấm tiếp các SV chưa chấm",
            command=lambda: self._start_ai_grading(only_ungraded=True),
        )
        self.btn_run_ungraded.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_load_existing = ttk.Button(
            action_frame,
            text="Mở Bảng Kiểm Duyệt Hiện Có (Tab 2)",
            command=self._load_and_switch_to_review,
        )
        self.btn_load_existing.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_stop_ai = ttk.Button(
            action_frame,
            text="Dừng Chấm",
            command=self._stop_ai_grading,
            state=tk.DISABLED,
        )
        self.btn_stop_ai.pack(side=tk.LEFT)

        prog_frame = ttk.Frame(self.tab1)
        prog_frame.pack(fill=tk.X, pady=(0, 6))
        self.ai_progress = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.ai_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.ai_status_var = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(prog_frame, textvariable=self.ai_status_var, width=32).pack(side=tk.RIGHT)

        log1_frame = ttk.LabelFrame(self.tab1, text="Nhật ký Kiểm tra Dữ liệu & Chấm AI", padding=6)
        log1_frame.pack(fill=tk.BOTH, expand=True)
        self.log1_text = scrolledtext.ScrolledText(log1_frame, height=12, font=("Consolas", 9), wrap=tk.WORD)
        self.log1_text.pack(fill=tk.BOTH, expand=True)

    # ==================== MÀN HÌNH CHÍNH (APPLE MINIMALIST SINGLE-SCREEN WORKSPACE) ====================
    def _build_tab2(self):
        # 1. Thanh điều khiển duy nhất sát đỉnh cửa sổ (y = 0, phong cách Apple macOS Toolbar, không Icon, không chữ AI)
        subbar = tk.Frame(
            self.tab2,
            bg="#FFFFFF",
            highlightbackground="#D2D2D7",
            highlightthickness=1,
            padx=12,
            pady=6,
        )
        subbar.pack(fill=tk.X)

        nav_left = tk.Frame(subbar, bg="#FFFFFF")
        nav_left.pack(side=tk.LEFT)

        self.btn_class_selector = tk.Button(
            nav_left,
            text="Lớp: (Chưa chọn)",
            bg="#1D1D1F",
            fg="#FFFFFF",
            activebackground="#3A3A3C",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._toggle_class_overlay,
        )
        self.btn_class_selector.pack(side=tk.LEFT, padx=(0, 8))

        seg_nav = tk.Frame(nav_left, bg="#E8E8ED", padx=2, pady=2)
        seg_nav.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_prev_sv = tk.Button(
            seg_nav,
            text="Trước",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._prev_student,
        )
        self.btn_prev_sv.pack(side=tk.LEFT, padx=(0, 2))

        self.btn_next_sv = tk.Button(
            seg_nav,
            text="Tiếp",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._next_student,
        )
        self.btn_next_sv.pack(side=tk.LEFT)

        self.active_sv_label = tk.Label(
            nav_left,
            text="Chưa nạp danh sách sinh viên",
            bg="#FFFFFF",
            fg="#1D1D1F",
            font=("Segoe UI", 10, "bold"),
        )
        self.active_sv_label.pack(side=tk.LEFT)

        self.active_sv_status_label = tk.Label(
            nav_left,
            text="",
            bg="#FFFFFF",
            fg="#6E6E73",
            font=("Segoe UI", 9),
            padx=6,
        )
        self.active_sv_status_label.pack(side=tk.LEFT)

        metrics_right = tk.Frame(subbar, bg="#FFFFFF")
        metrics_right.pack(side=tk.RIGHT)

        self.btn_toolbar_scraper = tk.Button(
            metrics_right,
            text="Hút dữ liệu",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._copy_js_scraper,
        )
        self.btn_toolbar_scraper.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_toolbar_ai = tk.Button(
            metrics_right,
            text="Chấm tự động",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._on_toolbar_ai_click,
        )
        self.btn_toolbar_ai.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_toolbar_pipeline = tk.Button(
            metrics_right,
            text="Tiến trình",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._toggle_pipeline_hud,
        )
        self.btn_toolbar_pipeline.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_toolbar_settings = tk.Button(
            metrics_right,
            text="Cài đặt",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._open_settings_and_log_dialog,
        )
        self.btn_toolbar_settings.pack(side=tk.LEFT, padx=(0, 6))

        # Cụm nút chỉnh cỡ chữ Segmented Control: A- | 11pt | A+
        font_seg = tk.Frame(metrics_right, bg="#E8E8ED", padx=2, pady=2)
        font_seg.pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            font_seg,
            text="A-",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
            cursor="hand2",
            command=lambda: self._change_font_size(-1),
        ).pack(side=tk.LEFT, padx=(0, 1))

        self.lbl_font_size = tk.Label(
            font_seg,
            text=f"{self.font_size_pt}pt",
            bg="#FFFFFF",
            fg="#1D1D1F",
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
        )
        self.lbl_font_size.pack(side=tk.LEFT, padx=(0, 1))

        tk.Button(
            font_seg,
            text="A+",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
            cursor="hand2",
            command=lambda: self._change_font_size(1),
        ).pack(side=tk.LEFT)

        self.btn_toggle_all_prompts = tk.Button(
            metrics_right,
            text="Ẩn cả 2 đề",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=7,
            pady=2,
            cursor="hand2",
            command=self._toggle_both_prompts,
        )
        self.btn_toggle_all_prompts.pack(side=tk.LEFT, padx=(0, 6))

        self.filter_mode_var = tk.StringVar(value="Tất cả")
        self.filter_combo = ttk.Combobox(
            metrics_right,
            textvariable=self.filter_mode_var,
            values=["Tất cả", "Chưa kiểm", "Đã kiểm", "Cần lưu ý"],
            state="readonly",
            width=9,
        )
        self.filter_combo.pack(side=tk.LEFT, padx=(0, 6))
        self.filter_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_review_filter())

        self.lbl_checked_prog = self._create_metric_badge(metrics_right, "Đã kiểm: 0 / 0", fg="#6E6E73")
        self.lbl_p1_avg = self._create_metric_badge(metrics_right, "P1: 0")
        self.lbl_p2_avg = self._create_metric_badge(metrics_right, "P2: 0")
        self.lbl_total_score = self._create_metric_badge(metrics_right, "Tổng: 0.0", fg="#1D1D1F")

        # 1B. Thanh Tiến Trình Hội Đồng Chấm Nổi (Floating Progress Bar — Dynamic Island style, 0px layout shift)
        self.pipeline_hud_frame = tk.Frame(
            self.tab2,
            bg="#FFFFFF",
            highlightbackground="#1D1D1F",
            highlightcolor="#0071E3",
            highlightthickness=1,
            padx=14,
            pady=10,
        )
        self._build_pipeline_hud_contents(self.pipeline_hud_frame)

        # 2. Không gian làm việc chính (Cột trái Apple Sidebar 236px + Khu vực chấm 2 tầng Part 1 & Part 2)
        workspace = tk.Frame(self.tab2, bg="#F5F5F7", padx=8, pady=8)
        workspace.pack(fill=tk.BOTH, expand=True)
        self.workspace_frame = workspace
        workspace.grid_rowconfigure(0, weight=1)
        workspace.grid_columnconfigure(0, weight=0, minsize=232)
        workspace.grid_columnconfigure(1, weight=1)

        sidebar = tk.Frame(
            workspace,
            bg="#FFFFFF",
            highlightbackground="#cbd5e1",
            highlightcolor="#2563eb",
            highlightthickness=2,
            width=236,
        )
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(1, weight=1)
        sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar = sidebar

        sb_header = tk.Frame(sidebar, bg="#FAFAFC", padx=8, pady=5)
        sb_header.grid(row=0, column=0, sticky="ew")
        tk.Label(
            sb_header,
            text="SINH VIÊN",
            bg="#FAFAFC",
            fg="#86868B",
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            sb_header,
            text="ĐIỂM / LƯU Ý",
            bg="#FAFAFC",
            fg="#86868B",
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.RIGHT)

        tree_frame = tk.Frame(sidebar, bg="#FFFFFF")
        tree_frame.grid(row=1, column=0, sticky="nsew")

        self.sv_tree = ttk.Treeview(
            tree_frame,
            columns=("tick", "name", "score", "note"),
            show="headings",
            selectmode="browse",
        )
        self.sv_tree.heading("tick", text="Duyệt")
        self.sv_tree.heading("name", text="Họ và tên")
        self.sv_tree.heading("score", text="Điểm")
        self.sv_tree.heading("note", text="Lưu ý")

        self.sv_tree.column("tick", width=36, minwidth=30, stretch=False, anchor=tk.CENTER)
        self.sv_tree.column("name", width=110, minwidth=85, stretch=True, anchor=tk.W)
        self.sv_tree.column("score", width=38, minwidth=34, stretch=False, anchor=tk.E)
        self.sv_tree.column("note", width=44, minwidth=38, stretch=False, anchor=tk.CENTER)

        self.sv_tree.tag_configure("checked", foreground="#1D1D1F")
        self.sv_tree.tag_configure("warning", foreground="#D70015")
        self.sv_tree.tag_configure("absent", foreground="#86868B")

        sv_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.sv_tree.yview)
        self.sv_tree.configure(yscrollcommand=sv_scroll.set)
        self.sv_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sv_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.sv_tree.bind("<<TreeviewSelect>>", self._on_sv_tree_select)
        self.sv_tree.bind("<Button-1>", self._on_sv_tree_click)

        sb_footer = tk.Frame(sidebar, bg="#FAFAFC", highlightbackground="#E5E5EA", highlightthickness=1, padx=8, pady=8)
        sb_footer.grid(row=2, column=0, sticky="ew")

        self.lbl_target_excel = tk.Label(
            sb_footer,
            text="File đích: (Chưa chọn file)",
            bg="#FAFAFC",
            fg="#6E6E73",
            font=("Segoe UI", 8),
            anchor="w",
            justify=tk.LEFT,
            wraplength=214,
        )
        self.lbl_target_excel.pack(fill=tk.X, pady=(0, 5))

        excel_btn_row = tk.Frame(sb_footer, bg="#FAFAFC")
        excel_btn_row.pack(fill=tk.X, pady=(0, 6))
        tk.Button(
            excel_btn_row,
            text="Đổi file",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=4,
            pady=3,
            cursor="hand2",
            command=self._change_target_excel_keep_scores,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        tk.Button(
            excel_btn_row,
            text="Nạp lại Excel",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=4,
            pady=3,
            cursor="hand2",
            command=self._reload_scores_from_excel,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 3))
        tk.Button(
            excel_btn_row,
            text="Mở",
            bg="#FFFFFF",
            fg="#6E6E73",
            activebackground="#F5F5F7",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=6,
            pady=3,
            cursor="hand2",
            command=self._open_excel_for_review,
        ).pack(side=tk.RIGHT)

        self.btn_approve_excel = tk.Button(
            sb_footer,
            text="Lưu điểm ra Excel",
            bg="#0071E3",
            fg="#FFFFFF",
            activebackground="#005BB5",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            pady=6,
            cursor="hand2",
            command=self._confirm_and_approve_to_excel,
        )
        self.btn_approve_excel.pack(fill=tk.X, pady=(0, 5))

        self.btn_sidebar_moodle_js = tk.Button(
            sb_footer,
            text="Sao chép mã nhập Moodle",
            bg="#1D1D1F",
            fg="#FFFFFF",
            activebackground="#3A3A3C",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            pady=6,
            cursor="hand2",
            command=self._copy_js_paced_grade_submitter,
        )
        self.btn_sidebar_moodle_js.pack(fill=tk.X)
        self._update_target_excel_label()

        # Cột phải: Khu vực đối chiếu 2 tầng cố định (Part 1 ở trên, Part 2 ở dưới)
        stage = tk.Frame(workspace, bg="#F5F5F7")
        stage.grid(row=0, column=1, sticky="nsew")
        stage.grid_rowconfigure(0, weight=1)
        stage.grid_rowconfigure(1, weight=1)
        stage.grid_columnconfigure(0, weight=1)

        self.part_widgets = {}
        self.part_widgets["p1"] = self._build_part_row(stage, part_key="p1", row_idx=0, title_prefix="Part 1")
        self.part_widgets["p2"] = self._build_part_row(stage, part_key="p2", row_idx=1, title_prefix="Part 2")

    def _create_metric_badge(self, parent, text: str, fg: str = "#1D1D1F") -> tk.Label:
        frame = tk.Frame(parent, bg="#F5F5F7", highlightbackground="#E5E5EA", highlightthickness=1, padx=8, pady=2)
        frame.pack(side=tk.LEFT, padx=(4, 0))
        lbl = tk.Label(frame, text=text, bg="#F5F5F7", fg=fg, font=("Segoe UI", 8, "bold"))
        lbl.pack()
        return lbl

    def _build_part_row(self, parent, part_key: str, row_idx: int, title_prefix: str) -> dict:
        card = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#D2D2D7", highlightthickness=1)
        card.grid(row=row_idx, column=0, sticky="nsew", pady=(0, 8) if row_idx == 0 else (0, 0))
        card.grid_rowconfigure(0, weight=1)
        card.grid_columnconfigure(0, weight=16, uniform=f"split_{part_key}")
        card.grid_columnconfigure(1, weight=5, minsize=240, uniform=f"split_{part_key}")

        # Nửa trái (~76%): Hiển thị Đề bài (in đậm câu hỏi) và Bài làm sinh viên song song (Side-by-Side)
        left_pane = tk.Frame(card, bg="#FFFFFF", highlightbackground="#E5E5EA", highlightthickness=1)
        left_pane.grid(row=0, column=0, sticky="nsew")
        left_pane.grid_rowconfigure(1, weight=1)
        left_pane.grid_columnconfigure(0, weight=1)

        header = tk.Frame(left_pane, bg="#FAFAFC", highlightbackground="#E5E5EA", highlightthickness=1, padx=12, pady=5)
        header.grid(row=0, column=0, sticky="ew")

        lbl_code = tk.Label(
            header,
            text=f"{title_prefix} · N/A",
            bg="#FAFAFC",
            fg="#1D1D1F",
            font=("Segoe UI", 9, "bold"),
        )
        lbl_code.pack(side=tk.LEFT)

        btn_prompt = tk.Button(
            header,
            text="Ẩn đề bài",
            bg="#FFFFFF",
            fg="#1D1D1F",
            activebackground="#F5F5F7",
            activeforeground="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=1,
            cursor="hand2",
            command=lambda pk=part_key: self._toggle_prompt_overlay(pk),
        )
        btn_prompt.pack(side=tk.RIGHT)

        lbl_meta = tk.Label(
            header,
            text="0 từ · 0 dòng",
            bg="#FAFAFC",
            fg="#86868B",
            font=("Segoe UI", 8),
            padx=10,
        )
        lbl_meta.pack(side=tk.RIGHT)

        body_container = tk.Frame(left_pane, bg="#FFFFFF")
        body_container.grid(row=1, column=0, sticky="nsew")
        body_container.grid_rowconfigure(0, weight=1)
        body_container.grid_columnconfigure(0, weight=9, uniform=f"pe_{part_key}")
        body_container.grid_columnconfigure(1, weight=13, uniform=f"pe_{part_key}")

        # Khung Đề bài bên trái (Nền #FAFAFC dịu mắt, tự động in đậm câu hỏi trọng tâm)
        prompt_pane = tk.Frame(body_container, bg="#FAFAFC", highlightbackground="#E5E5EA", highlightthickness=1)
        prompt_pane.grid(row=0, column=0, sticky="nsew")
        prompt_pane.grid_rowconfigure(0, weight=1)
        prompt_pane.grid_columnconfigure(0, weight=1)

        prompt_text = scrolledtext.ScrolledText(
            prompt_pane,
            font=("Segoe UI", max(9, self.font_size_pt - 1)),
            wrap=tk.WORD,
            bg="#FAFAFC",
            fg="#6E6E73",
            relief=tk.FLAT,
            padx=12,
            pady=8,
            spacing1=2,
            spacing3=3,
        )
        prompt_text.grid(row=0, column=0, sticky="nsew")
        prompt_text.tag_configure(
            "prompt_question",
            font=("Segoe UI", max(9, self.font_size_pt - 1), "bold"),
            foreground="#1D1D1F",
        )
        prompt_text.configure(state=tk.DISABLED)

        # Khung Bài làm sinh viên bên phải (Nền giấy trắng tinh #FFFFFF, chữ mực #1D1D1F to rõ)
        essay_pane = tk.Frame(body_container, bg="#FFFFFF", highlightbackground="#E5E5EA", highlightthickness=1)
        essay_pane.grid(row=0, column=1, sticky="nsew")
        essay_pane.grid_rowconfigure(0, weight=1)
        essay_pane.grid_columnconfigure(0, weight=1)

        essay_text = scrolledtext.ScrolledText(
            essay_pane,
            font=("Segoe UI", self.font_size_pt),
            wrap=tk.WORD,
            bg="#FFFFFF",
            fg="#1D1D1F",
            relief=tk.FLAT,
            padx=14,
            pady=8,
            spacing1=3,
            spacing3=3,
        )
        essay_text.grid(row=0, column=0, sticky="nsew")
        essay_text.configure(state=tk.DISABLED)

        # Nửa phải (~24%): Bảng chấm điểm & Nhận xét tối giản chuẩn Apple
        right_pane = tk.Frame(card, bg="#FFFFFF")
        right_pane.grid(row=0, column=1, sticky="nsew")
        right_pane.grid_columnconfigure(0, weight=0, minsize=26)
        right_pane.grid_columnconfigure(1, weight=0, minsize=34)
        right_pane.grid_columnconfigure(2, weight=1)
        for r_i in range(1, 5):
            right_pane.grid_rowconfigure(r_i, weight=1)

        hdr_frame_c0 = tk.Label(
            right_pane, text="MỤC", bg="#FAFAFC", fg="#86868B", font=("Segoe UI", 7, "bold"), anchor="center", padx=1, pady=5
        )
        hdr_frame_c0.grid(row=0, column=0, sticky="ew")
        hdr_frame_c1 = tk.Label(
            right_pane, text="ĐIỂM", bg="#FAFAFC", fg="#86868B", font=("Segoe UI", 7, "bold"), anchor="center", pady=5
        )
        hdr_frame_c1.grid(row=0, column=1, sticky="ew")
        hdr_frame_c2 = tk.Label(
            right_pane, text="NHẬN XÉT CHI TIẾT", bg="#FAFAFC", fg="#86868B", font=("Segoe UI", 7, "bold"), anchor="w", padx=6, pady=5
        )
        hdr_frame_c2.grid(row=0, column=2, sticky="ew")

        criteria_defs = [
            ("c", "C", "C"),
            ("o", "O", "O"),
            ("g", "G", "G"),
            ("v", "V", "V"),
        ]
        score_vars = {}
        score_entries = {}
        score_boxes = {}
        row_frames = {}
        reason_labels = {}

        for idx, (c_key, code_char, _short_lbl) in enumerate(criteria_defs, start=1):
            cell_crit = tk.Frame(right_pane, bg="#FFFFFF", highlightbackground="#F0F0F2", highlightthickness=1, padx=2)
            cell_crit.grid(row=idx, column=0, sticky="nsew")
            badge = tk.Label(
                cell_crit,
                text=code_char,
                bg="#F5F5F7",
                fg="#6E6E73",
                font=("Segoe UI", 8, "bold"),
                width=2,
                relief=tk.FLAT,
                bd=0,
            )
            badge.pack(expand=True, pady=1)

            cell_score = tk.Frame(right_pane, bg="#FFFFFF", highlightbackground="#F0F0F2", highlightthickness=1)
            cell_score.grid(row=idx, column=1, sticky="nsew")
            s_var = tk.StringVar(value="0")
            s_var.trace_add("write", lambda *_args, pk=part_key, ck=c_key: self._on_score_var_changed(pk, ck))

            score_box = tk.Frame(
                cell_score,
                bg="#FFFFFF",
                highlightbackground="#cbd5e1",
                highlightcolor="#2563eb",
                highlightthickness=2,
                padx=2,
                pady=1,
            )
            score_box.pack(expand=True, pady=1)

            entry = tk.Entry(
                score_box,
                textvariable=s_var,
                width=2,
                font=("Segoe UI", max(10, self.font_size_pt), "bold"),
                justify=tk.CENTER,
                relief=tk.FLAT,
                bd=0,
                bg="#FFFFFF",
                fg="#1D1D1F",
                insertbackground="#0071E3",
            )
            entry.pack()
            entry.bind("<FocusIn>", lambda _e, pk=part_key, ck=c_key: self._on_score_entry_focus_in(pk, ck))
            entry.bind("<FocusOut>", lambda _e, pk=part_key, ck=c_key: self._on_score_entry_focus_out(pk, ck))
            entry.bind("<KeyPress>", lambda e, pk=part_key, ck=c_key: self._on_score_entry_keypress(e, pk, ck))
            entry.bind("<Up>", self._on_key_up)
            entry.bind("<Down>", self._on_key_down)
            entry.bind("<Left>", self._on_key_left)
            entry.bind("<Right>", self._on_key_right)
            entry.bind("<space>", self._on_key_space)
            entry.bind("<Return>", lambda _e: self._focus_student_list())
            entry.bind("<Escape>", lambda _e: self._focus_student_list())

            score_vars[c_key] = s_var
            score_entries[c_key] = entry
            score_boxes[c_key] = score_box

            cell_reason = tk.Frame(right_pane, bg="#FFFFFF", highlightbackground="#F0F0F2", highlightthickness=1, padx=6, pady=2)
            cell_reason.grid(row=idx, column=2, sticky="nsew")
            lbl_reason = tk.Label(
                cell_reason,
                text="",
                bg="#FFFFFF",
                fg="#1D1D1F",
                font=("Segoe UI", max(9, self.font_size_pt - 2)),
                justify=tk.LEFT,
                anchor="w",
                wraplength=215,
            )
            lbl_reason.pack(fill=tk.BOTH, expand=True)
            cell_reason.bind(
                "<Configure>",
                lambda e, lbl=lbl_reason: lbl.configure(wraplength=max(110, e.width - 12)),
            )
            reason_labels[c_key] = lbl_reason
            row_frames[c_key] = (cell_crit, cell_score, cell_reason, badge, lbl_reason)

            for w_click in (cell_crit, badge, cell_score, score_box, cell_reason, lbl_reason):
                w_click.bind("<Button-1>", lambda _e, pk=part_key, ck=c_key: self._focus_score_by_part_crit(pk, ck))

        return {
            "lbl_code": lbl_code,
            "lbl_meta": lbl_meta,
            "btn_prompt": btn_prompt,
            "body_container": body_container,
            "prompt_pane": prompt_pane,
            "essay_pane": essay_pane,
            "essay_text": essay_text,
            "prompt_text": prompt_text,
            "score_vars": score_vars,
            "score_entries": score_entries,
            "score_boxes": score_boxes,
            "row_frames": row_frames,
            "reason_labels": reason_labels,
        }

    # ==================== TAB 3: TỰ ĐỘNG NHẬP ĐIỂM LÊN MOODLE ====================
    def _build_tab3(self):
        info_frame = ttk.LabelFrame(self.tab3, text="Bước 1: File Excel Phiếu Chấm Điểm Đã Approve", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 8))

        excel2_row = ttk.Frame(info_frame)
        excel2_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(excel2_row, text="File Excel Đã Approve:", width=20).pack(side=tk.LEFT)
        self.excel2_path_var = tk.StringVar(value=self.cfg.get("last_excel_path", ""))
        ttk.Entry(excel2_row, textvariable=self.excel2_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(excel2_row, text="Chọn File Excel...", command=self._browse_excel_tab2).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(excel2_row, text="Đọc & Làm mới Bảng", command=self._inspect_excel_tab2).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(excel2_row, text="Mở File Excel", command=self._open_excel_for_review).pack(side=tk.LEFT)

        self.excel_summary_var = tk.StringVar(value="Chưa tải thông tin file Excel.")
        ttk.Label(
            info_frame,
            textvariable=self.excel_summary_var,
            font=("Segoe UI", 9, "bold"),
            foreground="#1e3a8a",
        ).pack(anchor=tk.W)

        # Bảng xem trước điểm Q17 / Q18 sẽ nhập lên Moodle (Click vào SV nào để chọn chạy từ SV đó)
        preview_frame = ttk.LabelFrame(
            self.tab3,
            text="Bước 2: Bảng Xem Trước Điểm Sẽ Nhập Lên Moodle (Click chọn 1 dòng nếu muốn bắt đầu nhập từ sinh viên đó)",
            padding=8,
        )
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.web_tree = ttk.Treeview(
            preview_frame,
            columns=("row", "mssv", "ten", "q17", "q18", "status"),
            show="headings",
            height=8,
            selectmode="browse",
        )
        self.web_tree.heading("row", text="Dòng Excel")
        self.web_tree.heading("mssv", text="MSSV")
        self.web_tree.heading("ten", text="Họ và tên")
        self.web_tree.heading("q17", text="Điểm Q17 (Part 1)")
        self.web_tree.heading("q18", text="Điểm Q18 (Part 2)")
        self.web_tree.heading("status", text="Trạng thái")

        self.web_tree.column("row", width=80, stretch=False, anchor=tk.CENTER)
        self.web_tree.column("mssv", width=110, stretch=False, anchor=tk.CENTER)
        self.web_tree.column("ten", width=220, stretch=True, anchor=tk.W)
        self.web_tree.column("q17", width=120, stretch=False, anchor=tk.CENTER)
        self.web_tree.column("q18", width=120, stretch=False, anchor=tk.CENTER)
        self.web_tree.column("status", width=180, stretch=True, anchor=tk.W)

        web_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.web_tree.yview)
        self.web_tree.configure(yscrollcommand=web_scroll.set)
        self.web_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        web_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.web_tree.bind("<<TreeviewSelect>>", self._on_web_tree_select)

        opt_frame = ttk.LabelFrame(self.tab3, text="Bước 3: Cấu Hình Điều Hướng Tự Động Nhập Điểm Lên Trình Duyệt", padding=8)
        opt_frame.pack(fill=tk.X, pady=(0, 8))

        row_opt1 = ttk.Frame(opt_frame)
        row_opt1.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row_opt1, text="Dòng bắt đầu hoặc MSSV (Mặc định: 10):", width=34).pack(side=tk.LEFT)
        self.start_filter_var = tk.StringVar(value="10")
        ttk.Entry(row_opt1, textvariable=self.start_filter_var, width=16).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(row_opt1, text="Đặt lại từ đầu (Dòng 10)", command=lambda: self.start_filter_var.set("10")).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(row_opt1, text="Tốc độ nhập Web:").pack(side=tk.LEFT, padx=(0, 6))
        self.speed_var = tk.DoubleVar(value=1.0)
        ttk.Radiobutton(row_opt1, text="Nhanh (0.8x)", variable=self.speed_var, value=0.8).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(row_opt1, text="Chuẩn (1.0x)", variable=self.speed_var, value=1.0).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(row_opt1, text="Chậm (1.3x)", variable=self.speed_var, value=1.3).pack(side=tk.LEFT)

        btn2_frame = ttk.Frame(self.tab3)
        btn2_frame.pack(fill=tk.X, pady=(0, 6))

        self.btn_copy_js_grade = tk.Button(
            btn2_frame,
            text="Copy Mã JS Nhập Điểm Chạy Ngầm (F12 — Nhịp 3–5s/câu, Lớp đang chọn)",
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#1e40af",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._copy_js_paced_grade_submitter,
        )
        self.btn_copy_js_grade.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.btn_run_web = tk.Button(
            btn2_frame,
            text="Nhập Tự Động Click Chuột (Đếm ngược 5s)",
            bg="#0f172a",
            fg="#ffffff",
            activebackground="#1e293b",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._start_web_upload,
        )
        self.btn_run_web.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_stop_web = ttk.Button(
            btn2_frame,
            text="Dừng Nhập",
            command=self._stop_web_upload,
            state=tk.DISABLED,
        )
        self.btn_stop_web.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(
            btn2_frame,
            text="Mở Nhật Ký (nhat_ky_cham_diem.txt)",
            command=self._open_web_log_file,
        ).pack(side=tk.LEFT)

        prog2_frame = ttk.Frame(self.tab3)
        prog2_frame.pack(fill=tk.X, pady=(0, 6))
        self.web_progress = ttk.Progressbar(prog2_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.web_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.web_status_var = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(prog2_frame, textvariable=self.web_status_var, width=34).pack(side=tk.RIGHT)

        log2_frame = ttk.LabelFrame(self.tab3, text="Nhật ký Tự động Nhập Điểm Lên Web", padding=6)
        log2_frame.pack(fill=tk.BOTH, expand=True)
        self.log2_text = scrolledtext.ScrolledText(log2_frame, height=6, font=("Consolas", 9), wrap=tk.WORD)
        self.log2_text.pack(fill=tk.BOTH, expand=True)

    # ==================== KHỞI TẠO & PHÍM TẮT ====================
    def _bind_review_shortcuts(self):
        for seq, handler in (
            ("<Up>", self._on_key_up),
            ("<Down>", self._on_key_down),
            ("<Left>", self._on_key_left),
            ("<Right>", self._on_key_right),
            ("<space>", self._on_key_space),
            ("<Escape>", self._on_key_escape),
        ):
            self.root.bind(seq, handler)
            self.sv_tree.bind(seq, handler)

    def _is_review_tab_active(self) -> bool:
        return True

    def _is_score_entry_widget(self, widget) -> bool:
        if widget is None:
            return False
        for pk in ("p1", "p2"):
            entries = self.part_widgets.get(pk, {}).get("score_entries", {})
            if widget in entries.values():
                return True
        return False

    def _is_typing_in_input(self, event) -> bool:
        if event is None:
            return False
        w = getattr(event, "widget", None)
        if w is None:
            return False
        if self._is_score_entry_widget(w):
            return False
        cls_name = w.winfo_class()
        return cls_name in ("Entry", "TEntry", "Spinbox", "TSpinbox", "Text", "TCombobox")

    def _focus_student_list(self):
        self.focus_zone = "list"
        try:
            self.sv_tree.focus_set()
        except Exception:
            pass
        self._update_focus_visuals()

    def _focus_score_slot(self, slot_idx: int):
        slot_idx = max(0, min(len(self.SCORE_SLOTS) - 1, slot_idx))
        self.focus_zone = "scores"
        self.active_score_slot_idx = slot_idx
        pk, ck = self.SCORE_SLOTS[slot_idx]
        entry = self.part_widgets.get(pk, {}).get("score_entries", {}).get(ck)
        if entry is not None:
            try:
                entry.focus_set()
                entry.icursor(tk.END)
                entry.selection_clear()
            except Exception:
                pass
        self._update_focus_visuals()

    def _focus_score_by_part_crit(self, part_key: str, crit_key: str):
        if (part_key, crit_key) in self.SCORE_SLOTS:
            idx = self.SCORE_SLOTS.index((part_key, crit_key))
            self._focus_score_slot(idx)

    def _on_score_entry_focus_in(self, part_key: str, crit_key: str):
        if (part_key, crit_key) in self.SCORE_SLOTS:
            self.focus_zone = "scores"
            self.active_score_slot_idx = self.SCORE_SLOTS.index((part_key, crit_key))
            entry = self.part_widgets.get(part_key, {}).get("score_entries", {}).get(crit_key)
            if entry is not None:
                try:
                    entry.icursor(tk.END)
                    entry.selection_clear()
                except Exception:
                    pass
            self._update_focus_visuals()

    def _on_score_entry_focus_out(self, part_key: str, crit_key: str):
        if not self.filtered_indices:
            return
        s_var = self.part_widgets.get(part_key, {}).get("score_vars", {}).get(crit_key)
        if s_var is not None and not s_var.get().strip():
            real_idx = self.filtered_indices[self.current_review_pos]
            cur_val = self.review_items[real_idx].get(part_key, {}).get(crit_key, 0)
            self._updating_score_vars = True
            try:
                s_var.set(str(cur_val))
            finally:
                self._updating_score_vars = False

    def _on_score_entry_keypress(self, event, part_key: str, crit_key: str):
        keysym = getattr(event, "keysym", "")
        ch = getattr(event, "char", "")

        if keysym == "space" or ch == " ":
            self._toggle_check_and_advance()
            return "break"
        if keysym in ("Up", "Down", "Left", "Right", "Return", "Escape", "Tab"):
            return "break"
        if ch and ch in ("0", "1", "2", "3", "4", "5"):
            s_var = self.part_widgets[part_key]["score_vars"][crit_key]
            s_var.set(ch)
            entry = self.part_widgets[part_key]["score_entries"][crit_key]
            try:
                entry.icursor(tk.END)
                entry.selection_clear()
            except Exception:
                pass
            return "break"
        if keysym in ("BackSpace", "Delete"):
            self.part_widgets[part_key]["score_vars"][crit_key].set("")
            return "break"
        return "break"

    def _update_focus_visuals(self):
        active_pair = self.SCORE_SLOTS[self.active_score_slot_idx] if self.focus_zone == "scores" else None
        if hasattr(self, "sidebar"):
            if self.focus_zone == "list":
                self.sidebar.configure(highlightbackground="#2563eb", highlightcolor="#2563eb")
            else:
                self.sidebar.configure(highlightbackground="#cbd5e1", highlightcolor="#cbd5e1")

        for pk in ("p1", "p2"):
            w = self.part_widgets.get(pk, {})
            boxes = w.get("score_boxes", {})
            entries = w.get("score_entries", {})
            rows = w.get("row_frames", {})
            for ck in ("c", "o", "g", "v"):
                is_active = active_pair == (pk, ck)
                box = boxes.get(ck)
                ent = entries.get(ck)
                row_tuple = rows.get(ck)
                if box is not None:
                    box.configure(
                        highlightbackground="#2563eb" if is_active else "#cbd5e1",
                        highlightcolor="#2563eb",
                        bg="#dbeafe" if is_active else "#ffffff",
                    )
                if ent is not None:
                    ent.configure(
                        bg="#dbeafe" if is_active else "#ffffff",
                        fg="#1e3a8a" if is_active else "#0f172a",
                    )
                if row_tuple is not None:
                    cell_crit, cell_score, cell_reason, badge, lbl_reason = row_tuple
                    row_bg = "#eff6ff" if is_active else "#ffffff"
                    border_bg = "#bfdbfe" if is_active else "#f1f5f9"
                    cell_crit.configure(bg=row_bg, highlightbackground=border_bg)
                    cell_score.configure(bg=row_bg, highlightbackground=border_bg)
                    cell_reason.configure(bg=row_bg, highlightbackground=border_bg)
                    badge.configure(
                        bg="#1d4ed8" if is_active else "#eff6ff",
                        fg="#ffffff" if is_active else "#1d4ed8",
                    )
                    lbl_reason.configure(bg=row_bg)

    def _on_key_up(self, event=None):
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        if self.focus_zone == "scores":
            self._focus_score_slot(self.active_score_slot_idx - 1)
        else:
            self._prev_student()
        return "break"

    def _on_key_down(self, event=None):
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        if self.focus_zone == "scores":
            self._focus_score_slot(self.active_score_slot_idx + 1)
        else:
            self._next_student()
        return "break"

    def _on_key_left(self, event=None):
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        self._focus_student_list()
        return "break"

    def _on_key_right(self, event=None):
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        if self.focus_zone == "list":
            self._focus_score_slot(0)
        return "break"

    def _on_key_escape(self, event=None):
        if getattr(self, "class_overlay_visible", False):
            self._hide_class_overlay()
            return "break"
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        self._focus_student_list()
        return "break"

    def _on_key_space(self, event=None):
        if event is not None and (not self._is_review_tab_active() or self._is_typing_in_input(event)):
            return
        self._toggle_check_and_advance()
        return "break"

    def _initial_load(self):
        if self.active_class_code:
            active_cls = next((c for c in self.open_classes if c["code"] == self.active_class_code), None)
            if active_cls:
                j_p, e_p = resolve_or_recover_class_files(
                    self.active_class_code,
                    active_cls.get("json_path", ""),
                    active_cls.get("excel_path", ""),
                )
                active_cls["json_path"] = j_p
                active_cls["excel_path"] = e_p
                if j_p:
                    self.json_path_var.set(j_p)
                if e_p:
                    self.excel1_path_var.set(e_p)
                    self.excel2_path_var.set(e_p)
        json_p, excel_p = resolve_or_recover_class_files(
            self.active_class_code,
            self.json_path_var.get().strip(),
            self.excel1_path_var.get().strip(),
        )
        if json_p != self.json_path_var.get().strip():
            self.json_path_var.set(json_p)
        if excel_p != self.excel1_path_var.get().strip():
            self.excel1_path_var.set(excel_p)
            self.excel2_path_var.set(excel_p)

        self._update_target_excel_label()
        if excel_p and os.path.exists(excel_p):
            self._inspect_excel_tab2()
        if (json_p and os.path.exists(json_p)) or (excel_p and os.path.exists(excel_p)):
            self._refresh_data_matching(silent=True)
        else:
            self._sync_active_class_tab()

    # ==================== QUẢN LÝ ĐA LỚP HỌC & BỘ NHỚ ĐỘC LẬP TỪNG TAB ====================
    def _get_or_create_class_session(self, code: str) -> dict:
        if not code:
            code = "LOP_MOI"
        sess = self.class_sessions.get(code)
        if sess is None:
            cls_entry = next((c for c in self.open_classes if c["code"] == code), None)
            sess = {
                "code": code,
                "json_path": (cls_entry.get("json_path", "") if cls_entry else "").strip(),
                "excel_path": (cls_entry.get("excel_path", "") if cls_entry else "").strip(),
                "review_items": [],
                "current_review_pos": 0,
                "filter_mode": "Tất cả",
                "focus_zone": "list",
                "active_score_slot_idx": 0,
                "ai_running": False,
                "ai_queued": False,
                "ai_done": 0,
                "ai_total": 0,
                "ai_status": "Sẵn sàng.",
                "stop_ai_event": threading.Event(),
            }
            self.class_sessions[code] = sess
        return sess

    def _save_active_ui_to_class_session(self):
        if not self.active_class_code:
            return
        sess = self._get_or_create_class_session(self.active_class_code)
        sess["json_path"] = self.json_path_var.get().strip()
        sess["excel_path"] = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        if self.review_items:
            sess["review_items"] = self.review_items
        sess["current_review_pos"] = self.current_review_pos
        if hasattr(self, "filter_mode_var"):
            sess["filter_mode"] = self.filter_mode_var.get()
        sess["focus_zone"] = self.focus_zone
        sess["active_score_slot_idx"] = self.active_score_slot_idx

    def _sync_ai_controls_for_active_class(self):
        if not hasattr(self, "btn_run_ai"):
            return
        sess = self._get_or_create_class_session(self.active_class_code) if self.active_class_code else {}
        is_cls_running = bool(sess.get("ai_running"))
        is_cls_queued = bool(sess.get("ai_queued"))
        total = max(int(sess.get("ai_total") or 0), 1)
        done = int(sess.get("ai_done") or 0)
        self.ai_progress["maximum"] = total
        self.ai_progress["value"] = done if (is_cls_running or done > 0) else 0
        self.ai_status_var.set(str(sess.get("ai_status") or "Sẵn sàng."))

        if is_cls_running or is_cls_queued:
            self.btn_run_ai.config(state=tk.DISABLED)
            self.btn_run_ungraded.config(state=tk.DISABLED)
            self.btn_stop_ai.config(state=tk.NORMAL)
            if hasattr(self, "btn_toolbar_ai"):
                lbl = f"Dừng chấm ({done}/{total})" if is_cls_running and int(sess.get("ai_total") or 0) > 0 else "Dừng chấm"
                self.btn_toolbar_ai.configure(text=lbl, bg="#FFFBEB", fg="#B45309")
        else:
            self.btn_run_ai.config(state=tk.NORMAL)
            self.btn_run_ungraded.config(state=tk.NORMAL)
            self.btn_stop_ai.config(state=tk.DISABLED)
            if hasattr(self, "btn_toolbar_ai"):
                self.btn_toolbar_ai.configure(text="Chấm tự động", bg="#FFFFFF", fg="#1D1D1F")

    def _apply_class_bar_visibility(self):
        if not hasattr(self, "class_bar_frame") or not hasattr(self, "btn_toggle_class_bar"):
            return
        if self.class_bar_visible:
            self.class_bar_frame.pack(side=tk.LEFT, fill=tk.Y, pady=4, before=self.btn_toggle_class_bar)
            self.btn_toggle_class_bar.configure(
                text="Ẩn danh sách lớp",
                bg="#1D1D1F",
                fg="#FFFFFF",
            )
        else:
            self.class_bar_frame.pack_forget()
            self.btn_toggle_class_bar.configure(
                text=f"Lớp học ({len(self.open_classes)})",
                bg="#1D1D1F",
                fg="#FFFFFF",
            )

    def _toggle_class_bar(self):
        self.class_bar_visible = not self.class_bar_visible
        self._apply_class_bar_visibility()
        self._persist_class_bar_state()

    # ==================== FLOATING CLASS POPOVER & TYPOGRAPHY SCALING (APPLE DESIGN) ====================
    def _toggle_class_overlay(self):
        if getattr(self, "class_overlay_visible", False):
            self._hide_class_overlay()
        else:
            self._show_class_overlay()

    def _show_class_overlay(self):
        self.class_overlay_visible = True
        self._render_class_overlay_items()
        self.class_overlay_panel.place(x=12, y=42, width=410)
        self.class_overlay_panel.lift()

    def _hide_class_overlay(self):
        self.class_overlay_visible = False
        if hasattr(self, "class_overlay_panel"):
            self.class_overlay_panel.place_forget()

    def _on_root_click_close_overlay(self, event):
        if not getattr(self, "class_overlay_visible", False):
            return
        w = getattr(event, "widget", None)
        if w is None:
            return
        # Không đóng nếu click vào chính nút mở Popover hoặc các nút bên trong Popover
        curr = w
        while curr is not None:
            if curr is getattr(self, "class_overlay_panel", None):
                return
            if curr is getattr(self, "btn_class_selector", None) or curr is getattr(self, "btn_toggle_class_bar", None):
                return
            curr = getattr(curr, "master", None)
        self._hide_class_overlay()

    def _discover_workspace_classes(self) -> list[dict]:
        """Quét nhanh các file lớp học (.xlsx / bai_lam_iuh_*.json) có sẵn trong cùng thư mục để chuyển lớp 1-click."""
        open_codes = {c["code"].upper() for c in self.open_classes}
        search_dirs = []
        for c in self.open_classes:
            for k in ("excel_path", "json_path"):
                p = (c.get(k) or "").strip()
                if p and os.path.isdir(os.path.dirname(p)):
                    d = os.path.dirname(p)
                    if d not in search_dirs:
                        search_dirs.append(d)
        for fallback_d in (get_app_dir(), os.getcwd()):
            if fallback_d and os.path.isdir(fallback_d) and fallback_d not in search_dirs:
                search_dirs.append(fallback_d)

        found_by_code = {}
        for d in search_dirs:
            try:
                names = sorted(os.listdir(d))
            except Exception:
                continue
            for fn in names:
                low = fn.lower()
                if low.startswith("~$") or low.endswith("_review.json") or low == "config.json":
                    continue
                if not (low.endswith(".xlsx") or (low.endswith(".json") and "bai_lam" in low)):
                    continue
                full_p = os.path.join(d, fn)
                code = extract_explicit_iuh_class_code(full_p)
                if not code or code.upper() in open_codes or code.upper() in found_by_code:
                    continue
                found_by_code[code.upper()] = {"code": code, "path": full_p, "filename": fn}
                if len(found_by_code) >= 6:
                    break
        return list(found_by_code.values())

    def _render_class_overlay_items(self):
        if not hasattr(self, "class_overlay_panel"):
            return
        for child in self.class_overlay_panel.winfo_children():
            child.destroy()

        hdr = tk.Frame(self.class_overlay_panel, bg="#FFFFFF")
        hdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(
            hdr,
            text="DANH SÁCH LỚP ĐANG MỞ",
            bg="#FFFFFF",
            fg="#86868B",
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.LEFT)
        tk.Button(
            hdr,
            text="Đóng",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8),
            padx=8,
            pady=1,
            cursor="hand2",
            command=self._hide_class_overlay,
        ).pack(side=tk.RIGHT)

        if not self.open_classes:
            tk.Label(
                self.class_overlay_panel,
                text="Chưa có lớp nào được mở.",
                bg="#FFFFFF",
                fg="#6E6E73",
                font=("Segoe UI", 9),
                pady=6,
            ).pack(anchor="w")
        else:
            for cls_item in self.open_classes:
                code = cls_item["code"]
                is_active = code == self.active_class_code
                sess = self.class_sessions.get(code) or {}
                checked = int(cls_item.get("checked_count") or 0)
                total = int(cls_item.get("total_count") or 0)

                if sess.get("ai_running"):
                    ai_done = int(sess.get("ai_done") or 0)
                    ai_tot = int(sess.get("ai_total") or total or 0)
                    status_str = f"Đang chấm {ai_done}/{ai_tot}" if ai_tot > 0 else "Đang chấm..."
                elif sess.get("ai_queued"):
                    status_str = "Đang chờ hàng đợi"
                elif total > 0:
                    status_str = f"Đã duyệt {checked}/{total}"
                else:
                    status_str = "Chưa nạp dữ liệu"

                row_bg = "#F0F6FF" if is_active else "#FAFAFC"
                border_col = "#0071E3" if is_active else "#E5E5EA"
                row = tk.Frame(
                    self.class_overlay_panel,
                    bg=row_bg,
                    highlightbackground=border_col,
                    highlightthickness=1,
                    padx=8,
                    pady=5,
                )
                row.pack(fill=tk.X, pady=(0, 4))

                def _select_cls(c=code):
                    self._hide_class_overlay()
                    self._switch_to_class_tab(c)

                btn_main = tk.Button(
                    row,
                    text=f"{code}   ·   {status_str}",
                    bg=row_bg,
                    fg="#0071E3" if is_active else "#1D1D1F",
                    activebackground="#E5F0FF" if is_active else "#F5F5F7",
                    font=("Segoe UI", 9, "bold" if is_active else "normal"),
                    relief=tk.FLAT,
                    bd=0,
                    anchor="w",
                    cursor="hand2",
                    command=_select_cls,
                )
                btn_main.pack(side=tk.LEFT, fill=tk.X, expand=True)

                tk.Button(
                    row,
                    text="Đổi tên",
                    bg=row_bg,
                    fg="#6E6E73",
                    activebackground="#E8E8ED",
                    font=("Segoe UI", 8),
                    relief=tk.FLAT,
                    bd=0,
                    padx=6,
                    cursor="hand2",
                    command=lambda c=code: self._rename_class_tab_dialog(c),
                ).pack(side=tk.LEFT, padx=(4, 2))

                tk.Button(
                    row,
                    text="Đóng",
                    bg=row_bg,
                    fg="#6E6E73",
                    activebackground="#FEE2E2",
                    activeforeground="#D70015",
                    font=("Segoe UI", 8),
                    relief=tk.FLAT,
                    bd=0,
                    padx=6,
                    cursor="hand2",
                    command=lambda c=code: self._close_class_tab(c),
                ).pack(side=tk.RIGHT)

        discovered = self._discover_workspace_classes()
        if discovered:
            tk.Frame(self.class_overlay_panel, bg="#E5E5EA", height=1).pack(fill=tk.X, pady=6)
            tk.Label(
                self.class_overlay_panel,
                text="MỞ NHANH LỚP TRONG CÙNG THƯ MỤC",
                bg="#FFFFFF",
                fg="#86868B",
                font=("Segoe UI", 8, "bold"),
            ).pack(anchor="w", pady=(0, 4))
            for d_item in discovered:
                d_code = d_item["code"]
                d_path = d_item["path"]
                tk.Button(
                    self.class_overlay_panel,
                    text=f"{d_code}   ({d_item['filename']})",
                    bg="#FAFAFC",
                    fg="#1D1D1F",
                    activebackground="#F0F6FF",
                    font=("Segoe UI", 8),
                    relief=tk.SOLID,
                    bd=1,
                    anchor="w",
                    padx=8,
                    pady=3,
                    cursor="hand2",
                    command=lambda p=d_path: (self._hide_class_overlay(), self._open_or_activate_class_from_file(p)),
                ).pack(fill=tk.X, pady=(0, 3))

        tk.Frame(self.class_overlay_panel, bg="#E5E5EA", height=1).pack(fill=tk.X, pady=6)

        act_col = tk.Frame(self.class_overlay_panel, bg="#FFFFFF")
        act_col.pack(fill=tk.X)
        tk.Button(
            act_col,
            text="Mở thêm lớp học mới (Chọn JSON hoặc Excel)...",
            bg="#0071E3",
            fg="#FFFFFF",
            activebackground="#005BB5",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 8, "bold"),
            relief=tk.FLAT,
            pady=5,
            cursor="hand2",
            command=lambda: (self._hide_class_overlay(), self._add_class_tab_dialog()),
        ).pack(fill=tk.X, pady=(0, 4))

        sub_act = tk.Frame(act_col, bg="#FFFFFF")
        sub_act.pack(fill=tk.X)
        tk.Button(
            sub_act,
            text="Chọn File JSON lớp này...",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            font=("Segoe UI", 8),
            relief=tk.SOLID,
            bd=1,
            pady=3,
            cursor="hand2",
            command=lambda: (self._hide_class_overlay(), self._browse_json()),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        tk.Button(
            sub_act,
            text="Chọn File Excel lớp này...",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            font=("Segoe UI", 8),
            relief=tk.SOLID,
            bd=1,
            pady=3,
            cursor="hand2",
            command=lambda: (self._hide_class_overlay(), self._browse_excel_tab1()),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _change_font_size(self, delta: int):
        new_pt = max(9, min(16, int(getattr(self, "font_size_pt", 11)) + int(delta)))
        self.font_size_pt = new_pt
        save_config({"font_size_pt": new_pt})
        self._apply_font_size_to_widgets()

    def _apply_font_size_to_widgets(self):
        pt = int(getattr(self, "font_size_pt", 11))
        if hasattr(self, "lbl_font_size"):
            self.lbl_font_size.configure(text=f"{pt}pt")
        prompt_pt = max(9, pt - 1)
        essay_pt = pt
        reason_pt = max(9, pt - 1)
        score_pt = max(10, pt + 1)

        for pk in ("p1", "p2"):
            w = getattr(self, "part_widgets", {}).get(pk)
            if not w:
                continue
            pr_txt = w.get("prompt_text")
            if pr_txt is not None:
                pr_txt.configure(font=("Segoe UI", prompt_pt))
                pr_txt.tag_configure("prompt_question", font=("Segoe UI", prompt_pt, "bold"), foreground="#1D1D1F")
            es_txt = w.get("essay_text")
            if es_txt is not None:
                es_txt.configure(font=("Segoe UI", essay_pt))
            for ent in (w.get("score_entries") or {}).values():
                ent.configure(font=("Segoe UI", score_pt, "bold"))
            for lbl_r in (w.get("reason_labels") or {}).values():
                lbl_r.configure(font=("Segoe UI", reason_pt))

    def _build_pipeline_hud_contents(self, parent: tk.Frame):
        """Khởi tạo Thanh Tiến Trình Nổi (Floating Progress Bar) với Line tiến trình liền mạch 5 mốc."""
        self.pipeline_progress_pct = 0
        self.pipeline_detail_expanded = False

        hdr = tk.Frame(parent, bg="#FFFFFF")
        hdr.pack(fill=tk.X, pady=(0, 6))

        tk.Label(
            hdr,
            text="HỘI ĐỒNG CHẤM",
            bg="#FFFFFF",
            fg="#1D1D1F",
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 8))

        s_gk1 = _short_model_label(self.model_var.get())
        s_arb = _short_model_label(self.arbitrator_model_var.get())
        self.lbl_council_summary_pill = tk.Label(
            hdr,
            text=f"{s_gk1} ⇄ {s_arb}",
            bg="#F0F6FF",
            fg="#0071E3",
            font=("Segoe UI", 8, "bold"),
            padx=8,
            pady=1,
        )
        self.lbl_council_summary_pill.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_pipeline_pct = tk.Label(
            hdr,
            text="0%",
            bg="#F5F5F7",
            fg="#1D1D1F",
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
        )
        self.lbl_pipeline_pct.pack(side=tk.LEFT)

        right_ctrl = tk.Frame(hdr, bg="#FFFFFF")
        right_ctrl.pack(side=tk.RIGHT)

        chk_p2 = ttk.Checkbutton(
            right_ctrl,
            text="Thẩm định 2 vòng",
            variable=self.enable_pass2_var,
            command=self._on_council_config_changed,
        )
        chk_p2.pack(side=tk.LEFT, padx=(0, 6))

        tk.Label(right_ctrl, text="Vòng 2:", bg="#FFFFFF", fg="#6E6E73", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 3))
        cb_arb = ttk.Combobox(
            right_ctrl,
            textvariable=self.arbitrator_model_var,
            values=AVAILABLE_MODELS,
            state="readonly",
            width=16,
        )
        cb_arb.pack(side=tk.LEFT, padx=(0, 6))
        cb_arb.bind("<<ComboboxSelected>>", lambda _e: self._on_council_config_changed())

        self.btn_pipeline_detail_toggle = tk.Button(
            right_ctrl,
            text="Chi tiết",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8),
            padx=7,
            pady=1,
            cursor="hand2",
            command=self._toggle_pipeline_detail_drawer,
        )
        self.btn_pipeline_detail_toggle.pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(
            right_ctrl,
            text="✕",
            bg="#F5F5F7",
            fg="#1D1D1F",
            activebackground="#E8E8ED",
            relief=tk.FLAT,
            bd=0,
            font=("Segoe UI", 8, "bold"),
            padx=6,
            pady=1,
            cursor="hand2",
            command=self._hide_pipeline_hud,
        ).pack(side=tk.RIGHT)

        # Line tiến trình liền mạch (Continuous Progress Line trên tk.Canvas)
        self.pipeline_canvas = tk.Canvas(
            parent,
            height=52,
            bg="#FFFFFF",
            highlightthickness=0,
            bd=0,
        )
        self.pipeline_canvas.pack(fill=tk.X, pady=(0, 4))
        self.pipeline_canvas.bind("<Configure>", lambda _e: self._redraw_pipeline_canvas())

        step_defs = [
            (1, "Phân đề", "20 mã đề"),
            (2, "Chấm vòng 1", "4 cụm đề"),
            (3, "Đối chiếu chuẩn", "Quy tắc Khoa"),
            (4, "Thẩm định chéo", "Model vòng 2"),
            (5, "Hoàn tất", "Khớp dòng Excel"),
        ]
        self.pipeline_step_widgets = {}
        for idx, title_txt, sub_txt in step_defs:
            self.pipeline_step_widgets[idx] = {
                "state": "pending",
                "base_title": title_txt,
                "default_sub": sub_txt,
                "status_text": sub_txt,
            }

        # Dòng trạng thái trực tiếp 1 dòng (Single-line Live Status Ticker)
        ticker_row = tk.Frame(parent, bg="#FAFAFC", highlightbackground="#E5E5EA", highlightthickness=1, padx=8, pady=4)
        ticker_row.pack(fill=tk.X)
        self.lbl_pipeline_live_line = tk.Label(
            ticker_row,
            text="Sẵn sàng. Nhấn 'Chấm tự động' để chạy Hội đồng chấm 2 vòng.",
            bg="#FAFAFC",
            fg="#1D1D1F",
            font=("Segoe UI", 8),
            anchor="w",
        )
        self.lbl_pipeline_live_line.pack(fill=tk.X)

        # Ngăn xem nhật ký suy luận chi tiết (Mặc định thu gọn, bấm 'Chi tiết' để bung 4 dòng)
        self.pipeline_stream_box = tk.Frame(parent, bg="#FAFAFC", highlightbackground="#E5E5EA", highlightthickness=1, padx=6, pady=4)
        self.pipeline_stream_text = scrolledtext.ScrolledText(
            self.pipeline_stream_box,
            height=4,
            font=("Consolas", 8),
            bg="#FAFAFC",
            fg="#1D1D1F",
            relief=tk.FLAT,
            wrap=tk.WORD,
            padx=4,
            pady=2,
        )
        self.pipeline_stream_text.pack(fill=tk.X)
        self.pipeline_stream_text.insert(
            "1.0",
            "Sẵn sàng. Nhấn 'Chấm tự động' để xem trực tiếp Hội đồng chấm phân đề, suy luận và thẩm định chéo.\n",
        )
        self.pipeline_stream_text.configure(state=tk.DISABLED)
        self._redraw_pipeline_canvas()

    def _toggle_pipeline_detail_drawer(self):
        if not hasattr(self, "pipeline_stream_box"):
            return
        if self.pipeline_detail_expanded:
            self.pipeline_detail_expanded = False
            self.pipeline_stream_box.pack_forget()
            if hasattr(self, "btn_pipeline_detail_toggle"):
                self.btn_pipeline_detail_toggle.configure(bg="#F5F5F7", fg="#1D1D1F")
        else:
            self.pipeline_detail_expanded = True
            self.pipeline_stream_box.pack(fill=tk.X, pady=(4, 0))
            if hasattr(self, "btn_pipeline_detail_toggle"):
                self.btn_pipeline_detail_toggle.configure(bg="#1D1D1F", fg="#FFFFFF")

    def _redraw_pipeline_canvas(self):
        if not hasattr(self, "pipeline_canvas"):
            return
        cv = self.pipeline_canvas
        cv.delete("all")
        w = max(cv.winfo_width(), 560)
        pad_x = 56
        usable_w = max(w - 2 * pad_x, 200)
        y_line = 14
        xs = [pad_x + int(usable_w * i / 4.0) for i in range(5)]

        # 1. Đường nền xám nối từ mốc 1 đến mốc 5
        cv.create_line(xs[0], y_line, xs[-1], y_line, fill="#E5E5EA", width=3, capstyle=tk.ROUND)

        # 2. Đường line tiến trình актив chạy theo % thực tế
        pct = max(0, min(100, int(getattr(self, "pipeline_progress_pct", 0))))
        if pct > 0:
            active_end_x = xs[0] + int((xs[-1] - xs[0]) * (pct / 100.0))
            line_col = "#10B981" if pct >= 100 else "#0071E3"
            cv.create_line(xs[0], y_line, active_end_x, y_line, fill=line_col, width=3, capstyle=tk.ROUND)

        # 3. Vẽ 5 nút tròn + Nhãn chính + Nhãn phụ
        r = 8
        for i in range(5):
            idx = i + 1
            cx = xs[i]
            st_info = self.pipeline_step_widgets.get(idx, {})
            st = st_info.get("state", "pending")
            base_title = st_info.get("base_title", f"Bước {idx}")
            sub_txt = st_info.get("status_text", "")
            if len(sub_txt) > 24:
                sub_txt = sub_txt[:22] + "…"

            if st == "done":
                cv.create_oval(cx - r, y_line - r, cx + r, y_line + r, fill="#10B981", outline="#10B981", width=2)
                cv.create_text(cx, y_line, text="✓", fill="#FFFFFF", font=("Segoe UI", 8, "bold"))
                t_fg, s_fg = "#047857", "#065F46"
            elif st == "active":
                cv.create_oval(cx - r - 2, y_line - r - 2, cx + r + 2, y_line + r + 2, fill="#E5F0FF", outline="")
                cv.create_oval(cx - r, y_line - r, cx + r, y_line + r, fill="#0071E3", outline="#0071E3", width=2)
                cv.create_oval(cx - 3, y_line - 3, cx + 3, y_line + 3, fill="#FFFFFF", outline="")
                t_fg, s_fg = "#0071E3", "#1D1D1F"
            else:
                cv.create_oval(cx - r, y_line - r, cx + r, y_line + r, fill="#FFFFFF", outline="#D2D2D7", width=2)
                t_fg, s_fg = "#1D1D1F", "#86868B"

            cv.create_text(cx, y_line + 17, text=base_title, fill=t_fg, font=("Segoe UI", 8, "bold"))
            cv.create_text(cx, y_line + 30, text=sub_txt, fill=s_fg, font=("Segoe UI", 7))

    def _on_council_config_changed(self):
        gk1 = self.model_var.get().strip() or DEFAULT_MODEL
        arb = self.arbitrator_model_var.get().strip() or DEFAULT_ARBITRATOR_MODEL
        p2_on = bool(self.enable_pass2_var.get())
        save_config({
            "model_name": gk1,
            "arbitrator_model": arb,
            "enable_pass2_arbitrator": p2_on,
        })
        if hasattr(self, "lbl_council_summary_pill"):
            s_gk1 = _short_model_label(gk1)
            s_arb = _short_model_label(arb)
            if p2_on:
                self.lbl_council_summary_pill.configure(text=f"{s_gk1} ⇄ {s_arb}")
            else:
                self.lbl_council_summary_pill.configure(text=f"{s_gk1} · 1 vòng")

    def _show_pipeline_hud(self):
        if not hasattr(self, "pipeline_hud_frame"):
            return
        self._on_council_config_changed()
        if not self.pipeline_hud_visible:
            self.pipeline_hud_visible = True
            self.pipeline_hud_frame.place(relx=0.5, y=42, anchor="n", relwidth=0.78)
            self.pipeline_hud_frame.lift()
        else:
            self.pipeline_hud_frame.lift()
        self._redraw_pipeline_canvas()
        if hasattr(self, "btn_toolbar_pipeline"):
            self.btn_toolbar_pipeline.configure(bg="#1D1D1F", fg="#FFFFFF")

    def _hide_pipeline_hud(self):
        self.pipeline_hud_visible = False
        if hasattr(self, "pipeline_hud_frame"):
            self.pipeline_hud_frame.place_forget()
        if hasattr(self, "btn_toolbar_pipeline"):
            self.btn_toolbar_pipeline.configure(bg="#FFFFFF", fg="#1D1D1F")

    def _toggle_pipeline_hud(self):
        if self.pipeline_hud_visible:
            self._hide_pipeline_hud()
        else:
            self._show_pipeline_hud()

    def _reset_pipeline_hud_steps(self, initial_msg: str = ""):
        self.pipeline_progress_pct = 0
        if hasattr(self, "lbl_pipeline_pct"):
            self.lbl_pipeline_pct.configure(text="0%", bg="#F5F5F7", fg="#1D1D1F")
        for idx, w in self.pipeline_step_widgets.items():
            w["state"] = "pending"
            w["status_text"] = w["default_sub"]
        self._redraw_pipeline_canvas()
        if hasattr(self, "lbl_pipeline_live_line"):
            self.lbl_pipeline_live_line.configure(text=initial_msg or "Đang chuẩn bị tiến trình chấm...")
        if hasattr(self, "pipeline_stream_text"):
            self.pipeline_stream_text.configure(state=tk.NORMAL)
            self.pipeline_stream_text.delete("1.0", tk.END)
            if initial_msg:
                self.pipeline_stream_text.insert(tk.END, initial_msg + "\n")
            self.pipeline_stream_text.configure(state=tk.DISABLED)

    def _update_pipeline_hud_step(self, step_idx: int, state: str, badge_txt: str, detail_msg: str):
        w = self.pipeline_step_widgets.get(step_idx)
        if w is not None:
            w["state"] = state
            if badge_txt:
                w["status_text"] = badge_txt

            # Tính % đường line tiến trình nối liền mạch qua 5 mốc (0% -> 25% -> 50% -> 75% -> 100%)
            step_pct_map = {
                (1, "active"): 8,
                (1, "done"): 22,
                (2, "active"): 38,
                (2, "done"): 50,
                (3, "active"): 62,
                (3, "done"): 75,
                (4, "active"): 86,
                (4, "done"): 92,
                (5, "active"): 96,
                (5, "done"): 100,
            }
            new_pct = step_pct_map.get((step_idx, state))
            if new_pct is not None and new_pct >= getattr(self, "pipeline_progress_pct", 0):
                self.pipeline_progress_pct = new_pct
                if hasattr(self, "lbl_pipeline_pct"):
                    is_done_all = self.pipeline_progress_pct >= 100
                    self.lbl_pipeline_pct.configure(
                        text=f"{self.pipeline_progress_pct}%",
                        bg="#ECFDF5" if is_done_all else "#F0F6FF",
                        fg="#047857" if is_done_all else "#0071E3",
                    )
            self._redraw_pipeline_canvas()

        if detail_msg:
            if hasattr(self, "lbl_pipeline_live_line"):
                one_line = detail_msg.replace("\n", " ").strip()
                if len(one_line) > 118:
                    one_line = one_line[:116] + "…"
                self.lbl_pipeline_live_line.configure(text=one_line)
            if hasattr(self, "pipeline_stream_text"):
                self.pipeline_stream_text.configure(state=tk.NORMAL)
                self.pipeline_stream_text.insert(tk.END, detail_msg + "\n")
                self.pipeline_stream_text.see(tk.END)
                self.pipeline_stream_text.configure(state=tk.DISABLED)

    def _on_toolbar_ai_click(self):
        sess = self._get_or_create_class_session(self.active_class_code) if self.active_class_code else {}
        if sess.get("ai_running") or sess.get("ai_queued"):
            self._stop_ai_grading()
            return
        has_graded = any(it.get("graded") for it in self.review_items)
        has_ungraded = any(not it.get("graded") and not it.get("absent") for it in self.review_items)
        if has_graded and has_ungraded:
            ans = messagebox.askyesnocancel(
                "Chấm tự động",
                "Lớp này đã có một số bài được chấm.\n\n"
                "• Nhấn 'Yes' để chỉ chấm tiếp các sinh viên chưa có điểm.\n"
                "• Nhấn 'No' để chấm lại toàn bộ danh sách từ đầu.",
            )
            if ans is None:
                return
            self._start_ai_grading(only_ungraded=bool(ans))
        elif has_graded and not has_ungraded:
            if messagebox.askyesno(
                "Chấm lại toàn bộ",
                "Tất cả sinh viên có bài nộp trong lớp này đã có điểm.\nBạn có muốn chạy chấm tự động lại từ đầu không?",
            ):
                self._start_ai_grading(only_ungraded=False)
        else:
            self._start_ai_grading(only_ungraded=False)

    def _open_settings_and_log_dialog(self):
        if getattr(self, "_settings_win", None) is not None and self._settings_win.winfo_exists():
            self._settings_win.deiconify()
            self._settings_win.lift()
            self._settings_win.focus_force()
            return

        win = tk.Toplevel(self.root)
        win.title("Cài đặt hệ thống & Nhật ký")
        win.geometry("780x580")
        win.configure(bg="#F5F5F7")
        self._settings_win = win

        container = tk.Frame(win, bg="#F5F5F7", padx=14, pady=14)
        container.pack(fill=tk.BOTH, expand=True)

        card_cfg = tk.Frame(container, bg="#FFFFFF", highlightbackground="#D2D2D7", highlightthickness=1, padx=12, pady=10)
        card_cfg.pack(fill=tk.X, pady=(0, 10))

        tk.Label(card_cfg, text="CẤU HÌNH HỆ THỐNG & HỘI ĐỒNG ĐA MÔ HÌNH", bg="#FFFFFF", fg="#86868B", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        row1 = tk.Frame(card_cfg, bg="#FFFFFF")
        row1.pack(fill=tk.X, pady=(0, 6))
        tk.Label(row1, text="Khóa kết nối (API Key):", bg="#FFFFFF", fg="#1D1D1F", font=("Segoe UI", 9), width=18, anchor="w").pack(side=tk.LEFT)
        ent_key = ttk.Entry(row1, textvariable=self.api_key_var, show="*", width=32)
        ent_key.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        tk.Label(row1, text="GK1 (Vòng 1):", bg="#FFFFFF", fg="#1D1D1F", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        cb_model = ttk.Combobox(row1, textvariable=self.model_var, values=AVAILABLE_MODELS, state="readonly", width=18)
        cb_model.pack(side=tk.LEFT, padx=(0, 8))
        cb_model.bind("<<ComboboxSelected>>", lambda _e: self._on_council_config_changed())

        tk.Button(
            row1,
            text="Lưu cấu hình",
            bg="#0071E3",
            fg="#FFFFFF",
            activebackground="#005BB5",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 8, "bold"),
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._save_api_key,
        ).pack(side=tk.LEFT)

        row_council = tk.Frame(card_cfg, bg="#FFFFFF")
        row_council.pack(fill=tk.X, pady=(0, 6))
        tk.Label(row_council, text="Hội đồng Vòng 2:", bg="#FFFFFF", fg="#1D1D1F", font=("Segoe UI", 9), width=18, anchor="w").pack(side=tk.LEFT)
        ttk.Checkbutton(
            row_council,
            text="Bật Trọng tài Vòng 2 phản biện chéo các bài nhạy cảm/lệch điểm",
            variable=self.enable_pass2_var,
            command=self._on_council_config_changed,
        ).pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(row_council, text="Model Trọng tài:", bg="#FFFFFF", fg="#1D1D1F", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        cb_arb_dlg = ttk.Combobox(row_council, textvariable=self.arbitrator_model_var, values=AVAILABLE_MODELS, state="readonly", width=18)
        cb_arb_dlg.pack(side=tk.LEFT)
        cb_arb_dlg.bind("<<ComboboxSelected>>", lambda _e: self._on_council_config_changed())

        row2 = tk.Frame(card_cfg, bg="#FFFFFF")
        row2.pack(fill=tk.X, pady=(2, 0))
        tk.Label(row2, text="Bộ quy tắc chấm:", bg="#FFFFFF", fg="#1D1D1F", font=("Segoe UI", 9), width=18, anchor="w").pack(side=tk.LEFT)
        cb_rule = ttk.Combobox(
            row2,
            textvariable=self.rule_profile_label_var,
            values=[p["label"] for p in MARKING_RULE_PROFILES.values()],
            state="readonly",
            width=26,
        )
        cb_rule.pack(side=tk.LEFT, padx=(0, 8))
        cb_rule.bind("<<ComboboxSelected>>", self._on_rule_profile_selected)
        tk.Label(row2, textvariable=self.rules_info_var, bg="#FFFFFF", fg="#6E6E73", font=("Segoe UI", 8)).pack(side=tk.LEFT)
        tk.Button(
            row2,
            text="Nạp lại quy tắc",
            bg="#F5F5F7",
            fg="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._reset_rubric,
        ).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(
            row2,
            text="Mở file quy tắc",
            bg="#F5F5F7",
            fg="#1D1D1F",
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 8),
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._open_marking_rules_file,
        ).pack(side=tk.RIGHT)

        card_files = tk.Frame(container, bg="#FFFFFF", highlightbackground="#D2D2D7", highlightthickness=1, padx=12, pady=10)
        card_files.pack(fill=tk.X, pady=(0, 10))
        tk.Label(card_files, text="DỮ LIỆU LỚP HIỆN TẠI", bg="#FFFFFF", fg="#86868B", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))

        r_json = tk.Frame(card_files, bg="#FFFFFF")
        r_json.pack(fill=tk.X, pady=(0, 5))
        tk.Label(r_json, text="File JSON bài làm:", bg="#FFFFFF", fg="#1D1D1F", width=16, anchor="w").pack(side=tk.LEFT)
        ttk.Entry(r_json, textvariable=self.json_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r_json, text="Chọn JSON...", bg="#F5F5F7", fg="#1D1D1F", relief=tk.SOLID, bd=1, padx=8, command=self._browse_json).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(r_json, text="Kiểm tra đề", bg="#F5F5F7", fg="#1D1D1F", relief=tk.SOLID, bd=1, padx=8, command=self._check_json_info).pack(side=tk.LEFT)

        r_xls = tk.Frame(card_files, bg="#FFFFFF")
        r_xls.pack(fill=tk.X)
        tk.Label(r_xls, text="File Excel phiếu:", bg="#FFFFFF", fg="#1D1D1F", width=16, anchor="w").pack(side=tk.LEFT)
        ttk.Entry(r_xls, textvariable=self.excel1_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        tk.Button(r_xls, text="Chọn Excel...", bg="#F5F5F7", fg="#1D1D1F", relief=tk.SOLID, bd=1, padx=8, command=self._browse_excel_tab1).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(r_xls, text="Đối chiếu lại", bg="#F5F5F7", fg="#1D1D1F", relief=tk.SOLID, bd=1, padx=8, command=self._refresh_data_matching).pack(side=tk.LEFT)

        card_log = tk.Frame(container, bg="#FFFFFF", highlightbackground="#D2D2D7", highlightthickness=1, padx=12, pady=10)
        card_log.pack(fill=tk.BOTH, expand=True)
        tk.Label(card_log, text="NHẬT KÝ HOẠT ĐỘNG", bg="#FFFFFF", fg="#86868B", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 6))
        log_view = scrolledtext.ScrolledText(card_log, font=("Consolas", 9), bg="#FAFAFC", fg="#1D1D1F", relief=tk.FLAT, wrap=tk.WORD)
        log_view.pack(fill=tk.BOTH, expand=True)
        try:
            log_view.insert("1.0", self.log1_text.get("1.0", tk.END))
            log_view.see(tk.END)
        except Exception:
            pass

    def _persist_class_bar_state(self):
        save_config({
            "class_bar_visible": self.class_bar_visible,
            "open_classes": self.open_classes,
            "active_class_code": self.active_class_code,
        })

    def _sync_active_class_tab(self, auto_rename_by_code: bool = False):
        """Đồng bộ đường dẫn JSON/Excel và tiến độ duyệt của lớp đang mở vào bộ nhớ riêng và thanh Topbar."""
        json_p = self.json_path_var.get().strip()
        excel_p = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        if not json_p and not excel_p and not self.open_classes:
            self._render_class_bar()
            return

        detected_code = extract_class_code(json_p, excel_p)
        checked_cnt = sum(1 for it in self.review_items if it.get("checked")) if self.review_items else 0
        total_cnt = len(self.review_items) if self.review_items else 0

        active_cls = next((c for c in self.open_classes if c["code"] == self.active_class_code), None)
        if active_cls is None:
            new_code = detected_code
            active_cls = {
                "code": new_code,
                "json_path": json_p,
                "excel_path": excel_p,
                "checked_count": checked_cnt,
                "total_count": total_cnt,
            }
            self.open_classes.append(active_cls)
            self.active_class_code = new_code
        else:
            if json_p:
                active_cls["json_path"] = json_p
            if excel_p:
                active_cls["excel_path"] = excel_p
            if total_cnt > 0 or not (json_p or excel_p):
                active_cls["checked_count"] = checked_cnt
                active_cls["total_count"] = total_cnt
            if auto_rename_by_code and detected_code and detected_code != "LOP_MOI":
                if detected_code != active_cls["code"] and not any(
                    c["code"] == detected_code for c in self.open_classes if c is not active_cls
                ):
                    old_code = active_cls["code"]
                    active_cls["code"] = detected_code
                    self.active_class_code = detected_code
                    if old_code in self.class_sessions and detected_code not in self.class_sessions:
                        self.class_sessions[detected_code] = self.class_sessions.pop(old_code)
                        self.class_sessions[detected_code]["code"] = detected_code

        self._save_active_ui_to_class_session()
        self._persist_class_bar_state()
        self._apply_class_bar_visibility()
        self._render_class_bar()

    def _render_class_bar(self):
        # Cập nhật nút chọn lớp trên thanh công cụ chính (Apple Toolbar)
        if hasattr(self, "btn_class_selector"):
            if self.active_class_code:
                active_cls = next((c for c in self.open_classes if c["code"] == self.active_class_code), None)
                sess_act = self.class_sessions.get(self.active_class_code) or {}
                chk = sum(1 for it in self.review_items if it.get("checked")) if self.review_items else int((active_cls or {}).get("checked_count") or 0)
                tot = len(self.review_items) if self.review_items else int((active_cls or {}).get("total_count") or 0)
                if sess_act.get("ai_running"):
                    d_cnt = int(sess_act.get("ai_done") or 0)
                    t_cnt = int(sess_act.get("ai_total") or tot or 0)
                    self.btn_class_selector.configure(text=f"Lớp: {self.active_class_code} ({d_cnt}/{t_cnt})")
                elif tot > 0:
                    self.btn_class_selector.configure(text=f"Lớp: {self.active_class_code} ({chk}/{tot})")
                else:
                    self.btn_class_selector.configure(text=f"Lớp: {self.active_class_code}")
            else:
                self.btn_class_selector.configure(text="Lớp: Chọn lớp")

        if getattr(self, "class_overlay_visible", False):
            self._render_class_overlay_items()

        if not hasattr(self, "class_tabs_holder"):
            return
        for child in self.class_tabs_holder.winfo_children():
            child.destroy()

        if not self.open_classes:
            tk.Label(
                self.class_tabs_holder,
                text="(Chưa mở lớp)",
                bg="#0f172a",
                fg="#64748b",
                font=("Segoe UI", 8, "italic"),
            ).pack(side=tk.LEFT, padx=(0, 4))
            return

        for cls_item in self.open_classes:
            code = cls_item["code"]
            is_active = code == self.active_class_code
            sess = self.class_sessions.get(code) or {}
            checked = int(cls_item.get("checked_count") or 0)
            total = int(cls_item.get("total_count") or 0)

            if sess.get("ai_running"):
                ai_done = int(sess.get("ai_done") or 0)
                ai_tot = int(sess.get("ai_total") or total or 0)
                prog_txt = f" ({ai_done}/{ai_tot})" if ai_tot > 0 else " (Đang chấm...)"
            elif sess.get("ai_queued"):
                prog_txt = " (Chờ...)"
            else:
                prog_txt = f" ({checked}/{total})" if total > 0 else ""

            if is_active:
                pill_bg = "#2563eb"
                pill_fg = "#ffffff"
                border_col = "#60a5fa"
            elif sess.get("ai_running"):
                pill_bg = "#78350f"
                pill_fg = "#fde68a"
                border_col = "#f59e0b"
            else:
                pill_bg = "#1e293b"
                pill_fg = "#cbd5e1"
                border_col = "#334155"

            pill = tk.Frame(
                self.class_tabs_holder,
                bg=pill_bg,
                highlightbackground=border_col,
                highlightcolor=border_col,
                highlightthickness=1,
            )
            pill.pack(side=tk.LEFT, padx=(0, 4))

            btn_tab = tk.Button(
                pill,
                text=f"{code}{prog_txt}",
                bg=pill_bg,
                fg=pill_fg,
                activebackground="#1d4ed8" if is_active else "#334155",
                activeforeground="#ffffff",
                font=("Segoe UI", 8, "bold" if (is_active or sess.get("ai_running")) else "normal"),
                relief=tk.FLAT,
                bd=0,
                padx=7,
                pady=1,
                cursor="hand2",
                command=lambda c=code: self._switch_to_class_tab(c),
            )
            btn_tab.pack(side=tk.LEFT)
            btn_tab.bind("<Double-Button-1>", lambda _e, c=code: self._rename_class_tab_dialog(c))

            btn_close = tk.Button(
                pill,
                text="×",
                bg=pill_bg,
                fg="#bfdbfe" if is_active else "#94a3b8",
                activebackground="#dc2626",
                activeforeground="#ffffff",
                font=("Segoe UI", 8, "bold"),
                relief=tk.FLAT,
                bd=0,
                padx=4,
                pady=1,
                cursor="hand2",
                command=lambda c=code: self._close_class_tab(c),
            )
            btn_close.pack(side=tk.LEFT)

    def _clear_review_ui_when_incomplete(self):
        self.review_items = []
        self.filtered_indices = []
        self.current_review_pos = 0
        for child in self.sv_tree.get_children():
            self.sv_tree.delete(child)
        self.lbl_checked_prog.configure(text="Đã kiểm: 0 / 0")
        self.active_sv_label.configure(text="Chưa đủ dữ liệu của lớp này — Hãy chọn file JSON hoặc Excel")
        self.active_sv_status_label.configure(text="")
        self.match_summary_var.set(
            "Chưa tìm thấy dữ liệu của lớp này. Vui lòng chọn File JSON hoặc File Excel ở Bước 2."
        )

    def _switch_to_class_tab(self, target_code: str):
        if self.is_running_web:
            messagebox.showwarning("Đang nhập điểm web", "Vui lòng đợi tiến trình nhập điểm Moodle hoàn tất trước khi chuyển lớp.")
            return
        target_cls = next((c for c in self.open_classes if c["code"] == target_code), None)
        if not target_cls:
            return

        # 1. Lưu toàn bộ bộ nhớ phiên (điểm, dòng đang chọn, ô điểm đang gõ, bộ lọc) của lớp hiện tại trước khi chuyển
        if self.active_class_code and self.review_items:
            self._auto_save_review_session()
            self._save_active_ui_to_class_session()

        # 2. Chuyển sang lớp đích và tự động khôi phục đường dẫn file nếu bị thiếu 1 trong 2 file
        self.active_class_code = target_code
        sess = self._get_or_create_class_session(target_code)
        raw_json = target_cls.get("json_path", "").strip() or sess.get("json_path", "").strip()
        raw_excel = target_cls.get("excel_path", "").strip() or sess.get("excel_path", "").strip()
        json_p, excel_p = resolve_or_recover_class_files(target_code, raw_json, raw_excel)
        target_cls["json_path"] = json_p
        target_cls["excel_path"] = excel_p
        sess["json_path"] = json_p
        sess["excel_path"] = excel_p

        self.json_path_var.set(json_p)
        self.excel1_path_var.set(excel_p)
        self.excel2_path_var.set(excel_p)
        save_config({
            "last_json_path": json_p,
            "last_excel_path": excel_p,
            "active_class_code": target_code,
        })

        self._update_target_excel_label()
        self._sync_ai_controls_for_active_class()

        if excel_p and os.path.exists(excel_p):
            self._inspect_excel_tab2()
        else:
            for child in self.web_tree.get_children():
                self.web_tree.delete(child)
            self.excel_summary_var.set(f"Lớp {target_code}: Chưa chọn file Excel.")

        # 3. Nếu bộ nhớ tạm của tab lớp này đã có sẵn danh sách SV -> Khôi phục tức thì (0ms, giữ nguyên vị trí SV đang dò)
        if sess.get("review_items"):
            self.review_items = sess["review_items"]
            if hasattr(self, "filter_mode_var") and sess.get("filter_mode"):
                self.filter_mode_var.set(sess["filter_mode"])
            self.current_review_pos = int(sess.get("current_review_pos") or 0)
            self.focus_zone = str(sess.get("focus_zone") or "list")
            self.active_score_slot_idx = int(sess.get("active_score_slot_idx") or 0)
            self._update_match_summary_from_items(self.review_items)
            self._apply_review_filter()
            self._update_focus_visuals()
            self._render_class_bar()
            return

        # 4. Nếu chưa có trong RAM -> Nạp từ đĩa (hỗ trợ cả khi chỉ có 1 file Excel hoặc 1 file JSON + _review.json)
        if (json_p and os.path.exists(json_p)) or (excel_p and os.path.exists(excel_p)):
            self.current_review_pos = 0
            self._refresh_data_matching(silent=True)
        else:
            try:
                recovered_items = build_review_items_from_excel_and_json("", "", class_code=target_code)
            except Exception:
                recovered_items = []
            if recovered_items:
                self.review_items = recovered_items
                sess["review_items"] = recovered_items
                self.current_review_pos = 0
                self._update_match_summary_from_items(recovered_items)
                self._sync_active_class_tab()
                self._apply_review_filter()
            else:
                self._clear_review_ui_when_incomplete()
                self._render_class_bar()

    def _add_class_tab_dialog(self):
        if self.is_running_web:
            messagebox.showwarning("Đang nhập điểm web", "Vui lòng đợi tác vụ nhập điểm web hoàn tất trước khi mở thêm lớp.")
            return
        selected = filedialog.askopenfilename(
            title="Mở thêm Lớp học mới — Chọn file JSON bài làm hoặc file Excel Phiếu chấm của lớp",
            filetypes=[
                ("Dữ liệu Lớp học (JSON hoặc Excel)", "*.json *.xlsx"),
                ("JSON bài làm (bai_lam_iuh_*.json)", "*.json"),
                ("Excel Phiếu chấm (*.xlsx)", "*.xlsx"),
            ],
        )
        if not selected:
            return
        self._open_or_activate_class_from_file(selected)

    def _open_or_activate_class_from_file(self, selected_path: str):
        """Mở tab lớp mới hoặc chuyển sang tab lớp tương ứng từ 1 file (JSON hoặc Excel) mà không làm lẫn lộn lớp cũ."""
        if self.active_class_code and self.review_items:
            self._auto_save_review_session()
            self._save_active_ui_to_class_session()

        json_p, excel_p = auto_pair_class_files(selected_path)
        base_code = extract_class_code(json_p, excel_p)
        json_p, excel_p = resolve_or_recover_class_files(base_code, json_p, excel_p)

        existing = next(
            (
                c
                for c in self.open_classes
                if c["code"].upper() == base_code.upper()
                or (json_p and os.path.normcase(c.get("json_path", "")) == os.path.normcase(json_p))
                or (excel_p and os.path.normcase(c.get("excel_path", "")) == os.path.normcase(excel_p))
            ),
            None,
        )
        if existing is not None:
            if json_p:
                existing["json_path"] = json_p
            if excel_p:
                existing["excel_path"] = excel_p
            self._switch_to_class_tab(existing["code"])
            return

        new_code = base_code
        suffix = 2
        while any(c["code"].upper() == new_code.upper() for c in self.open_classes):
            new_code = f"{base_code}_{suffix}"
            suffix += 1

        new_entry = {
            "code": new_code,
            "json_path": json_p,
            "excel_path": excel_p,
            "checked_count": 0,
            "total_count": 0,
        }
        self.open_classes.append(new_entry)
        if not self.class_bar_visible:
            self.class_bar_visible = True
            self._apply_class_bar_visibility()
        self._switch_to_class_tab(new_code)

    def _close_class_tab(self, target_code: str):
        sess = self.class_sessions.get(target_code) or {}
        if sess.get("ai_running"):
            messagebox.showwarning("Lớp đang chấm AI", f"Lớp {target_code} đang được AI chấm ngầm. Vui lòng đợi hoặc nhấn Dừng trước khi đóng tab này.")
            return
        idx = next((i for i, c in enumerate(self.open_classes) if c["code"] == target_code), None)
        if idx is None:
            return
        if target_code == self.active_class_code:
            self._auto_save_review_session()
            self._save_active_ui_to_class_session()

        self.open_classes.pop(idx)
        self.ai_queue = [(c, u) for (c, u) in self.ai_queue if c != target_code]
        if not self.open_classes:
            self.active_class_code = ""
            self._persist_class_bar_state()
            self._apply_class_bar_visibility()
            self._render_class_bar()
            return

        if target_code == self.active_class_code:
            next_idx = min(idx, len(self.open_classes) - 1)
            self._switch_to_class_tab(self.open_classes[next_idx]["code"])
        else:
            self._persist_class_bar_state()
            self._apply_class_bar_visibility()
            self._render_class_bar()

    def _rename_class_tab_dialog(self, target_code: str):
        cls_item = next((c for c in self.open_classes if c["code"] == target_code), None)
        if not cls_item:
            return
        new_name = simpledialog.askstring(
            "Đổi tên Tab Lớp học",
            f"Nhập tên/mã lớp hiển thị trên thanh tab cho '{target_code}':",
            initialvalue=target_code,
            parent=self.root,
        )
        if not new_name or not new_name.strip():
            return
        clean_name = new_name.strip()
        if clean_name != target_code and any(c["code"].lower() == clean_name.lower() for c in self.open_classes):
            messagebox.showwarning("Trùng tên lớp", f"Đã có tab lớp mang tên '{clean_name}'.")
            return
        cls_item["code"] = clean_name
        if target_code in self.class_sessions:
            self.class_sessions[clean_name] = self.class_sessions.pop(target_code)
            self.class_sessions[clean_name]["code"] = clean_name
        if self.active_class_code == target_code:
            self.active_class_code = clean_name
        self._persist_class_bar_state()
        self._render_class_bar()

    def _route_selected_file_safely(self, selected_path: str, is_json: bool) -> bool:
        """
        Chốt chặn chống lẫn lộn file gốc giữa các lớp:
        Nếu đang đứng ở Tab lớp A (VD: 2607AVCB005) mà chọn nhầm file của lớp B (VD: 2607AVCB004),
        tự động mở/chuyển sang Tab riêng cho lớp B thay vì ghi đè lên đường dẫn của lớp A!
        Trả về True nếu đã tự động chuyển sang Tab lớp mới.
        """
        new_explicit = extract_explicit_iuh_class_code(selected_path)
        cur_json = self.json_path_var.get().strip()
        cur_excel = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        cur_explicit = (
            extract_explicit_iuh_class_code(self.active_class_code)
            or extract_explicit_iuh_class_code(cur_json)
            or extract_explicit_iuh_class_code(cur_excel)
        )
        if new_explicit and cur_explicit and new_explicit != cur_explicit:
            self._open_or_activate_class_from_file(selected_path)
            return True

        # Nếu tab hiện tại mới chỉ có 1 file, tự động ghép cặp file còn lại cùng lớp (nếu có trong cùng thư mục)
        paired_j, paired_e = auto_pair_class_files(selected_path)
        if is_json:
            self.json_path_var.set(selected_path)
            if not cur_excel and paired_e:
                self.excel1_path_var.set(paired_e)
                self.excel2_path_var.set(paired_e)
        else:
            self.excel1_path_var.set(selected_path)
            self.excel2_path_var.set(selected_path)
            if not cur_json and paired_j:
                self.json_path_var.set(paired_j)
        return False

    # ==================== TIỆN ÍCH GIAO DIỆN ====================
    def _update_target_excel_label(self):
        if not hasattr(self, "lbl_target_excel"):
            return
        excel_p = self.excel1_path_var.get().strip() or (
            self.excel2_path_var.get().strip() if hasattr(self, "excel2_path_var") else ""
        )
        if excel_p:
            self.lbl_target_excel.configure(
                text=f"Excel đích: {os.path.basename(excel_p)}",
                fg="#1e3a8a" if os.path.exists(excel_p) else "#b91c1c",
            )
        else:
            self.lbl_target_excel.configure(text="Excel đích: (Chưa chọn file)", fg="#b91c1c")

    def _change_target_excel_keep_scores(self) -> str:
        """Đổi file Excel đích ngay tại Tab 2 mà KHÔNG làm mất điểm vừa nhập trong Suite và KHÔNG trộn nhầm lớp khác."""
        cur_p = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        init_dir = os.path.dirname(cur_p) if cur_p and os.path.exists(os.path.dirname(cur_p)) else ""
        path = filedialog.askopenfilename(
            title="Chọn File Excel Phiếu Chấm Đích để Xuất Điểm (.xlsx) — Giữ nguyên điểm đã chấm",
            initialdir=init_dir or None,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return ""

        # Kiểm tra chống chọn nhầm file Excel của lớp khác khi đang giữ bảng điểm của lớp hiện tại
        cur_json = self.json_path_var.get().strip()
        try:
            validate_class_files_not_mixed(cur_json, path)
        except ValueError as err_mix:
            messagebox.showerror("Cảnh báo nhầm file Excel khác lớp", str(err_mix))
            return ""

        self.excel1_path_var.set(path)
        self.excel2_path_var.set(path)
        save_config({"last_excel_path": path})
        self._update_target_excel_label()

        # Cập nhật lại số dòng Excel (row) theo MSSV trên file Excel đích mới nhưng giữ nguyên 100% điểm & trạng thái checked trong self.review_items
        if self.review_items and os.path.exists(path):
            try:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                all_rows = list(ws.iter_rows(values_only=True))
                wb.close()
                cols = detect_excel_schema(all_rows)
                c_mssv = cols.get("mssv")
                if c_mssv is not None:
                    row_map = {}
                    for r_idx, r in enumerate(all_rows, 1):
                        if r and c_mssv < len(r):
                            m = normalize_mssv(r[c_mssv])
                            if m:
                                row_map[m] = r_idx
                    matched_cnt = sum(1 for it in self.review_items if normalize_mssv(it.get("mssv")) in row_map)
                    if row_map and matched_cnt == 0:
                        messagebox.showerror(
                            "Cảnh báo lệch danh sách sinh viên",
                            f"File Excel '{os.path.basename(path)}' không có bất kỳ MSSV nào trùng với lớp đang chấm!\n"
                            f"Hệ thống đã chặn để bảo vệ file gốc khỏi bị ghi đè nhầm lớp.",
                        )
                        return ""
                    for it in self.review_items:
                        m = normalize_mssv(it.get("mssv"))
                        if m in row_map:
                            it["row"] = row_map[m]
                self._auto_save_review_session()
                self._render_review_detail()
            except Exception:
                pass
        self._sync_active_class_tab(auto_rename_by_code=True)
        self._inspect_excel_tab2()
        return path

    def _log1(self, msg: str):
        self.log1_text.insert(tk.END, msg + "\n")
        self.log1_text.see(tk.END)

    def _log2(self, msg: str):
        self.log2_text.insert(tk.END, msg + "\n")
        self.log2_text.see(tk.END)

    def _toggle_api_key_visibility(self):
        self.api_entry.config(show="" if self.show_key_var.get() else "*")

    def _save_api_key(self):
        key = self.api_key_var.get().strip()
        model = self.model_var.get().strip() or DEFAULT_MODEL
        arb_model = self.arbitrator_model_var.get().strip() or DEFAULT_ARBITRATOR_MODEL
        p2_on = bool(self.enable_pass2_var.get())
        save_config({
            "api_key": key,
            "model_name": model,
            "arbitrator_model": arb_model,
            "enable_pass2_arbitrator": p2_on,
        })
        self._on_council_config_changed()
        messagebox.showinfo(
            "Đã lưu",
            f"Đã lưu cấu hình Hội đồng AI:\n"
            f"• Giám khảo 1 (Vòng 1): {model}\n"
            f"• Trọng tài Vòng 2: {arb_model} ({'Đang bật' if p2_on else 'Đang tắt'})\n"
            f"• File cấu hình: {get_config_path()}",
        )

    def _format_rules_info_text(self) -> str:
        prof = MARKING_RULE_PROFILES.get(self.rule_profile, MARKING_RULE_PROFILES[DEFAULT_RULE_PROFILE])
        return f"File: {prof['filename']} ({len(self.rubric_content.splitlines())} dòng C-O-G-V)"

    def _on_rule_profile_selected(self, _event=None):
        chosen_label = self.rule_profile_label_var.get().strip()
        for k, p in MARKING_RULE_PROFILES.items():
            if p["label"] == chosen_label:
                self.rule_profile = k
                break
        self.rubric_content = load_marking_rules(self.rule_profile)
        self.rules_info_var.set(self._format_rules_info_text())
        save_config({
            "rule_profile": self.rule_profile,
            "rubric_text": self.rubric_content,
        })
        prof = MARKING_RULE_PROFILES[self.rule_profile]
        self._log1(f"Đã chuyển sang bộ quy tắc chấm: {prof['label']} ({prof['filename']}).")

    def _reset_rubric(self):
        self.rubric_content = load_marking_rules(self.rule_profile)
        self.rules_info_var.set(self._format_rules_info_text())
        save_config({
            "rule_profile": self.rule_profile,
            "rubric_text": self.rubric_content,
        })
        prof = MARKING_RULE_PROFILES[self.rule_profile]
        self._log1(f"Đã nạp lại bộ quy tắc chấm: {prof['label']} ({prof['filename']}).")

    def _open_marking_rules_file(self):
        rules_path = get_marking_rules_path(self.rule_profile)
        if not os.path.exists(rules_path):
            load_marking_rules(self.rule_profile)
        try:
            os.startfile(rules_path)
        except Exception as e:
            messagebox.showerror("Lỗi mở file", f"Không thể mở file {rules_path}:\n{e}")

    def _copy_js_scraper(self):
        pyperclip.copy(JS_SCRAPER_CODE)
        self._log1("Đã copy đoạn mã JS Hút Bài (8 luồng, lấy cả Chữ + Ảnh đề + Định dạng dòng & tự đặt tên theo Mã lớp) vào Clipboard.")
        messagebox.showinfo(
            "Đã Copy Mã Hút Bài (F12)",
            "Đã copy mã JS 8 luồng vào Clipboard.\n\n"
            "Hướng dẫn thực hiện trên trình duyệt (Edge/Chrome):\n"
            "1. Mở trang Bảng điểm của lớp trên luyenthinn.iuh.edu.vn\n"
            "2. Nhấn phím F12 (hoặc Ctrl + Shift + I) -> chuyển sang thẻ Console\n"
            "3. Dán (Ctrl + V) và nhấn Enter.\n"
            "4. Đợi vài giây, trình duyệt sẽ tự tải file 'bai_lam_iuh_<MãLớp>.json' về máy.",
        )

    def _browse_json(self):
        path = filedialog.askopenfilename(
            title="Chọn file bai_lam_iuh.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            if self._route_selected_file_safely(path, is_json=True):
                self._check_json_info()
                return
            save_config({"last_json_path": self.json_path_var.get().strip()})
            self._check_json_info()
            self._sync_active_class_tab(auto_rename_by_code=True)
            self._refresh_data_matching(silent=True)

    def _check_json_info(self):
        path = self.json_path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Chưa có file JSON", "Vui lòng chọn file bai_lam_iuh.json hợp lệ.")
            return
        try:
            info = inspect_json_dataset(path)
            self._log1("------------------------------------------------------------------")
            self._log1(f"BÁO CÁO KIỂM TRA ĐỀ THI & BÀI LÀM TRONG: {os.path.basename(path)}")
            self._log1(f"  - Tổng số sinh viên trong file JSON: {info['total']} (Có nộp bài viết: {info['non_empty_students']})")
            self._log1(
                f"  - Part 1 (Q.17): Lấy được đề của {info['p1_with_prompt']} SV | "
                f"Số Mã đề: {len(info['p1_codes'])} ({', '.join(info['p1_codes'][:6])}) | "
                f"Ảnh đề đính kèm: {info['p1_img_count']}"
            )
            self._log1(
                f"  - Part 2 (Q.18): Lấy được đề của {info['p2_with_prompt']} SV | "
                f"Số Mã đề: {len(info['p2_codes'])} ({', '.join(info['p2_codes'][:6])}) | "
                f"Ảnh đề đính kèm: {info['p2_img_count']}"
            )
            self._log1("------------------------------------------------------------------")
        except Exception as e:
            messagebox.showerror("Lỗi đọc file JSON", f"Không thể đọc file JSON:\n{e}")

    def _browse_excel_tab1(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel Phiếu Chấm Writing (.xlsx)",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            if self._route_selected_file_safely(path, is_json=False):
                return
            save_config({"last_excel_path": self.excel1_path_var.get().strip()})
            self._update_target_excel_label()
            self._sync_active_class_tab(auto_rename_by_code=True)
            self._inspect_excel_tab2()
            self._refresh_data_matching(silent=True)

    def _browse_excel_tab2(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel Phiếu Chấm Đã Approve (.xlsx)",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            if self._route_selected_file_safely(path, is_json=False):
                return
            save_config({"last_excel_path": self.excel2_path_var.get().strip()})
            self._update_target_excel_label()
            self._sync_active_class_tab(auto_rename_by_code=True)
            self._inspect_excel_tab2()
            self._refresh_data_matching(silent=True)

    def _update_match_summary_from_items(self, items: list[dict]):
        total_sv = len(items)
        present_sv = sum(1 for it in items if not it.get("absent"))
        absent_sv = total_sv - present_sv
        graded_sv = sum(1 for it in items if it.get("graded"))
        checked_sv = sum(1 for it in items if it.get("checked"))
        self.match_summary_var.set(
            f"Đã nạp: {total_sv} SV  |  Có bài nộp: {present_sv} SV  |  "
            f"Vắng thi: {absent_sv} SV  |  Đã có điểm chấm: {graded_sv}/{present_sv} SV  |  Đã tick kiểm: {checked_sv}/{total_sv}"
        )

    def _refresh_data_matching(self, silent: bool = False):
        raw_json = self.json_path_var.get().strip()
        raw_excel = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        json_path, excel_path = resolve_or_recover_class_files(self.active_class_code, raw_json, raw_excel)
        if json_path != raw_json:
            self.json_path_var.set(json_path)
        if excel_path != raw_excel:
            self.excel1_path_var.set(excel_path)
            self.excel2_path_var.set(excel_path)

        self._update_target_excel_label()
        has_json = bool(json_path and os.path.exists(json_path))
        has_excel = bool(excel_path and os.path.exists(excel_path))

        if not has_json and not has_excel:
            self._sync_active_class_tab(auto_rename_by_code=True)
            if not silent:
                messagebox.showwarning("Thiếu file", "Vui lòng chọn File JSON bài làm hoặc File Excel Phiếu chấm của lớp.")
            return
        try:
            items = build_review_items_from_excel_and_json(
                excel_path if has_excel else "",
                json_path if has_json else "",
                class_code=self.active_class_code,
            )
            self.review_items = items
            if self.active_class_code:
                sess = self._get_or_create_class_session(self.active_class_code)
                sess["review_items"] = items
                sess["json_path"] = json_path
                sess["excel_path"] = excel_path
            self._update_match_summary_from_items(items)
            self._sync_active_class_tab(auto_rename_by_code=True)
            self._apply_review_filter()
        except Exception as e:
            if not silent:
                messagebox.showerror("Lỗi đối chiếu dữ liệu", f"Chi tiết lỗi:\n{e}")

    def _load_and_switch_to_review(self):
        self._refresh_data_matching(silent=False)

    # ==================== LOGIC TAB 2: KIỂM DUYỆT & SỬA ĐIỂM ====================
    @staticmethod
    def _calc_item_scores(item: dict) -> tuple[float, float, float]:
        p1 = item.get("p1") or {}
        p2 = item.get("p2") or {}
        raw1 = int(p1.get("c", 0)) + int(p1.get("o", 0)) + int(p1.get("g", 0)) + int(p1.get("v", 0))
        raw2 = int(p2.get("c", 0)) + int(p2.get("o", 0)) + int(p2.get("g", 0)) + int(p2.get("v", 0))
        f1 = raw1 / 4.0
        f2 = raw2 / 4.0
        # Khớp 100% với công thức =MROUND(I+O, 0.5) của Excel (làm tròn .25 và .75 lên 0.5 tiếp theo, không dùng làm tròn chẵn của Python)
        tot = ((raw1 + raw2 + 1) // 2) / 2.0
        return f1, f2, tot

    def _apply_review_filter(self):
        mode = self.filter_mode_var.get()
        indices = []
        for idx, it in enumerate(self.review_items):
            if mode == "Chưa kiểm" and it.get("checked"):
                continue
            if mode == "Đã kiểm" and not it.get("checked"):
                continue
            if mode == "Cần lưu ý" and not it.get("note"):
                continue
            indices.append(idx)
        self.filtered_indices = indices
        if self.current_review_pos >= len(self.filtered_indices):
            self.current_review_pos = max(0, len(self.filtered_indices) - 1)
        self._render_review_sidebar()
        self._render_review_detail()

    def _render_review_sidebar(self):
        for child in self.sv_tree.get_children():
            self.sv_tree.delete(child)

        checked_total = sum(1 for it in self.review_items if it.get("checked"))
        self.lbl_checked_prog.configure(text=f"Đã kiểm: {checked_total} / {len(self.review_items)}")

        for pos, real_idx in enumerate(self.filtered_indices):
            it = self.review_items[real_idx]
            _, _, tot = self._calc_item_scores(it)
            tick_str = "✓" if it.get("checked") else "☐"
            stt_str = f"{real_idx + 1:02d}. {it.get('ten', '')}"
            score_str = "-" if it.get("absent") else f"{tot:.1f}"
            note_str = it.get("note") or ""

            tags = []
            if it.get("absent"):
                tags.append("absent")
            elif it.get("checked"):
                tags.append("checked")
            elif note_str:
                tags.append("warning")

            iid = str(pos)
            self.sv_tree.insert("", tk.END, iid=iid, values=(tick_str, stt_str, score_str, note_str), tags=tuple(tags))

        if self.filtered_indices and 0 <= self.current_review_pos < len(self.filtered_indices):
            sel_iid = str(self.current_review_pos)
            if self.sv_tree.exists(sel_iid):
                self.sv_tree.selection_set(sel_iid)
                self.sv_tree.see(sel_iid)

    @staticmethod
    def _reason_text_color(reason_txt: str) -> str:
        low = (reason_txt or "").lower()
        if any(k in low for k in ("lỗi", "lạc đề", "khóa", "thiếu", "sai ", "nửa đúng", "quá cụt", "bỏ trống", "chép lại", "yếu")):
            return "#D70015"
        return "#1D1D1F"

    @staticmethod
    def _insert_highlighted_prompt(pr_widget: scrolledtext.ScrolledText, text: str):
        """Tự động in đậm các câu hỏi / gạch đầu dòng yêu cầu chính trong đề bài để giảng viên liếc mắt nắm ngay ý chính."""
        lines = (text or "").splitlines(keepends=True)
        if not lines:
            pr_widget.insert("1.0", "[Không có nội dung đề bài dạng chữ]")
            return
        question_starters = (
            "-", "•", "*", "1.", "2.", "3.", "4.", "a.", "b.", "c.",
            "write ", "in your ", "your email ", "your letter ", "your essay ",
            "discuss ", "explain ", "describe ", "answer ", "include ",
            "what ", "why ", "how ", "where ", "when ", "who ", "which ", "do you ", "should ",
        )
        for line in lines:
            stripped = line.strip()
            low = stripped.lower()
            is_key_line = (
                "?" in stripped
                or any(low.startswith(st) for st in question_starters)
            )
            if is_key_line and stripped:
                pr_widget.insert(tk.END, line, ("prompt_question",))
            else:
                pr_widget.insert(tk.END, line)

    def _render_review_detail(self):
        if not self.filtered_indices:
            self.active_sv_label.configure(text="Không có sinh viên nào trong bộ lọc hiện tại")
            self.active_sv_status_label.configure(text="")
            return

        real_idx = self.filtered_indices[self.current_review_pos]
        it = self.review_items[real_idx]
        f1, f2, tot = self._calc_item_scores(it)

        self.active_sv_label.configure(
            text=f"{real_idx + 1:02d}. {it.get('ten', '')}   ·   MSSV: {it.get('mssv', '')}   ·   Dòng {it.get('row', '')}"
        )
        c_badge = str(it.get("council_badge") or "").strip()
        if it.get("checked"):
            status_txt = f"Đã duyệt   ·   {c_badge}" if c_badge else "Đã duyệt"
            status_fg = "#0071E3" if "2 vòng" in c_badge else "#1D1D1F"
        else:
            status_txt = f"Chưa duyệt   ·   {c_badge}" if c_badge else "Chưa duyệt"
            status_fg = "#0071E3" if "2 vòng" in c_badge else "#86868B"
        self.active_sv_status_label.configure(text=status_txt, fg=status_fg)

        self.lbl_p1_avg.configure(text=f"P1: {f1:g}")
        self.lbl_p2_avg.configure(text=f"P2: {f2:g}")
        self.lbl_total_score.configure(text=f"Cộng: {f1 + f2:g} → Tổng: {tot:.1f}")

        self._updating_score_vars = True
        try:
            for pk, prefix in (("p1", "Part 1"), ("p2", "Part 2")):
                p_data = it.get(pk) or {}
                w = self.part_widgets[pk]
                w["lbl_code"].configure(text=f"{prefix} · Đề {p_data.get('code', 'N/A')}")
                w["lbl_meta"].configure(text=f"Bài làm: {p_data.get('word_count', 0)} từ · {p_data.get('line_count', 0)} dòng")

                # Hiển thị bài làm nguyên bản như sinh viên nộp trên Moodle
                essay_widget: scrolledtext.ScrolledText = w["essay_text"]
                essay_widget.configure(state=tk.NORMAL)
                essay_widget.delete("1.0", tk.END)
                essay_raw = p_data.get("essay") or "[BỎ TRỐNG]"
                essay_widget.insert("1.0", essay_raw)
                essay_widget.configure(state=tk.DISABLED)

                # Nội dung đề bài (đã chuẩn hóa format dòng + in đậm câu hỏi trọng tâm) hiển thị song song với bài làm
                pr_widget: scrolledtext.ScrolledText = w["prompt_text"]
                pr_widget.configure(state=tk.NORMAL)
                pr_widget.delete("1.0", tk.END)
                pr_raw = clean_prompt_text(p_data.get("prompt") or "")
                imgs = p_data.get("images") or []
                if imgs:
                    pr_raw = f"[Đề bài có {len(imgs)} hình ảnh đính kèm]\n\n" + pr_raw
                self._insert_highlighted_prompt(pr_widget, pr_raw or "[Không có nội dung đề bài dạng chữ]")
                pr_widget.configure(state=tk.DISABLED)

                for c_key, r_key in (("c", "rc"), ("o", "ro"), ("g", "rg"), ("v", "rv")):
                    w["score_vars"][c_key].set(str(p_data.get(c_key, 0)))
                    reason_txt = str(p_data.get(r_key, "") or "")
                    lbl_r: tk.Label = w["reason_labels"][c_key]
                    lbl_r.configure(
                        text=reason_txt,
                        fg=self._reason_text_color(reason_txt),
                    )
        finally:
            self._updating_score_vars = False

    def _apply_prompt_visibility_ui(self, part_key: str):
        w = self.part_widgets[part_key]
        body: tk.Frame = w["body_container"]
        pr_pane: tk.Frame = w["prompt_pane"]
        btn: tk.Button = w["btn_prompt"]
        if self.prompt_visible[part_key]:
            pr_pane.grid(row=0, column=0, sticky="nsew")
            body.grid_columnconfigure(0, weight=9, uniform=f"pe_{part_key}")
            body.grid_columnconfigure(1, weight=13, uniform=f"pe_{part_key}")
            btn.configure(text="Ẩn đề bài", bg="#FFFFFF", fg="#1D1D1F")
        else:
            pr_pane.grid_remove()
            body.grid_columnconfigure(0, weight=0, uniform="")
            body.grid_columnconfigure(1, weight=1, uniform="")
            btn.configure(text="Hiện đề bài", bg="#1D1D1F", fg="#FFFFFF")
        if hasattr(self, "btn_toggle_all_prompts"):
            both_shown = self.prompt_visible.get("p1", True) and self.prompt_visible.get("p2", True)
            self.btn_toggle_all_prompts.configure(text="Ẩn cả 2 đề" if both_shown else "Hiện cả 2 đề")

    def _toggle_prompt_overlay(self, part_key: str):
        self.prompt_visible[part_key] = not self.prompt_visible[part_key]
        self._apply_prompt_visibility_ui(part_key)

    def _toggle_both_prompts(self):
        target = not (self.prompt_visible.get("p1", True) and self.prompt_visible.get("p2", True))
        for pk in ("p1", "p2"):
            self.prompt_visible[pk] = target
            self._apply_prompt_visibility_ui(pk)

    def _on_score_var_changed(self, part_key: str, crit_key: str):
        if self._updating_score_vars or not self.filtered_indices:
            return
        raw_val = self.part_widgets[part_key]["score_vars"][crit_key].get().strip()
        if not raw_val:
            return
        try:
            num = max(0, min(5, int(float(raw_val))))
        except ValueError:
            return

        real_idx = self.filtered_indices[self.current_review_pos]
        it = self.review_items[real_idx]
        if it.get("absent"):
            return
        p_data = it[part_key]
        p_data[crit_key] = num
        it["graded"] = True

        # Nếu giảng viên gõ Content = 0 cho bài lạc đề (>= 12 từ), tự động áp dụng quy tắc 0 - 1 - 1 - 1
        w_widgets = self.part_widgets[part_key]
        is_c0_lock = crit_key == "c" and num == 0 and int(p_data.get("word_count", 0)) >= 12
        if is_c0_lock:
            p_data["o"] = 1
            p_data["g"] = 1
            p_data["v"] = 1
            self._updating_score_vars = True
            try:
                for ck_lock in ("o", "g", "v"):
                    w_widgets["score_vars"][ck_lock].set("1")
            finally:
                self._updating_score_vars = False

        # Chỉ làm mới lại lý do của đúng tiêu chí mà Giám khảo vừa sửa điểm (trừ khi khóa C=0 0-1-1-1),
        # giữ nguyên nhận xét chi tiết của 3 tiêu chí còn lại để không làm mất dẫn chứng AI đã chấm.
        rc, ro, rg, rv = _synthesize_part_cogv_reasons(
            p_data.get("prompt") or "",
            p_data.get("essay") or "",
            int(p_data.get("c", 0)),
            int(p_data.get("o", 0)),
            int(p_data.get("g", 0)),
            int(p_data.get("v", 0)),
            "" if crit_key == "c" else (p_data.get("rc") or ""),
            "" if crit_key == "o" else (p_data.get("ro") or ""),
            "" if crit_key == "g" else (p_data.get("rg") or ""),
            "" if crit_key == "v" else (p_data.get("rv") or ""),
            is_part1=(part_key == "p1"),
            format_type=p_data.get("format_type") or "BINH_THUONG",
            force_rebuild=is_c0_lock,
        )
        p_data["rc"], p_data["ro"], p_data["rg"], p_data["rv"] = rc, ro, rg, rv
        for ck_r, r_txt in (("c", rc), ("o", ro), ("g", rg), ("v", rv)):
            lbl_r: tk.Label = w_widgets["reason_labels"][ck_r]
            lbl_r.configure(text=r_txt, fg=self._reason_text_color(r_txt))

        # Cập nhật tag cảnh báo ở sidebar nếu đổi thành Lạc đề (Content=0)
        p1_c = int((it.get("p1") or {}).get("c", 0))
        p2_c = int((it.get("p2") or {}).get("c", 0))
        if p1_c == 0 or p2_c == 0:
            if not it.get("note"):
                it["note"] = "Lạc đề"
        elif it.get("note") == "Lạc đề" and p1_c >= 3 and p2_c >= 3:
            it["note"] = ""

        f1, f2, tot = self._calc_item_scores(it)
        self.lbl_p1_avg.configure(text=f"P1: {f1:g}")
        self.lbl_p2_avg.configure(text=f"P2: {f2:g}")
        self.lbl_total_score.configure(text=f"Cộng: {f1 + f2:g} → Tổng: {tot:.1f}")

        # Cập nhật điểm & lưu ý trên dòng danh sách bên trái ngay lập tức
        iid = str(self.current_review_pos)
        if self.sv_tree.exists(iid):
            vals = list(self.sv_tree.item(iid, "values"))
            if len(vals) >= 4:
                vals[2] = f"{tot:.1f}"
                vals[3] = it.get("note") or ""
                self.sv_tree.item(iid, values=vals)

        self._auto_save_review_session()

    def _auto_save_review_session(self):
        json_p = self.json_path_var.get().strip()
        excel_p = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        if (json_p or excel_p or self.active_class_code) and self.review_items:
            try:
                save_review_session(json_p, excel_p, self.review_items, class_code=self.active_class_code)
            except Exception:
                pass
        self._sync_active_class_tab()

    def _on_sv_tree_select(self, _event):
        sel = self.sv_tree.selection()
        if not sel:
            return
        try:
            pos = int(sel[0])
        except ValueError:
            return
        if 0 <= pos < len(self.filtered_indices) and pos != self.current_review_pos:
            self.current_review_pos = pos
            self._render_review_detail()

    def _on_sv_tree_click(self, event):
        self.focus_zone = "list"
        self._update_focus_visuals()
        region = self.sv_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.sv_tree.identify_column(event.x)
        row_id = self.sv_tree.identify_row(event.y)
        if col == "#1" and row_id:
            try:
                pos = int(row_id)
            except ValueError:
                return
            if 0 <= pos < len(self.filtered_indices):
                real_idx = self.filtered_indices[pos]
                self.review_items[real_idx]["checked"] = not self.review_items[real_idx].get("checked", False)
                self.current_review_pos = pos
                self._auto_save_review_session()
                self._render_review_sidebar()
                self._render_review_detail()
                return "break"

    def _prev_student(self):
        if self.filtered_indices and self.current_review_pos > 0:
            self.current_review_pos -= 1
            self._render_review_sidebar()
            self._render_review_detail()

    def _next_student(self):
        if self.filtered_indices and self.current_review_pos < len(self.filtered_indices) - 1:
            self.current_review_pos += 1
            self._render_review_sidebar()
            self._render_review_detail()

    def _toggle_check_and_advance(self):
        if not self.filtered_indices:
            return
        real_idx = self.filtered_indices[self.current_review_pos]
        self.review_items[real_idx]["checked"] = True
        self._auto_save_review_session()
        if self.current_review_pos < len(self.filtered_indices) - 1:
            self.current_review_pos += 1
        self._focus_student_list()
        self._render_review_sidebar()
        self._render_review_detail()

    def _confirm_and_approve_to_excel(self):
        if not self.review_items:
            messagebox.showwarning("Chưa có dữ liệu duyệt", "Chưa có danh sách điểm sinh viên để xuất ra Excel.")
            return
        excel_path = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        if not excel_path or not os.path.exists(excel_path):
            excel_path = self._change_target_excel_keep_scores()
            if not excel_path or not os.path.exists(excel_path):
                return

        self._auto_save_review_session()
        total_sv = len(self.review_items)
        checked_sv = sum(1 for it in self.review_items if it.get("checked"))
        unchecked_sv = total_sv - checked_sv

        warn_line = (
            f"\nLưu ý: Còn {unchecked_sv} sinh viên chưa được đánh dấu Tick kiểm duyệt.\n"
            if unchecked_sv > 0
            else "\nTất cả sinh viên trong danh sách đã được đánh dấu kiểm duyệt.\n"
        )
        confirm_msg = (
            f"Bạn đã kiểm duyệt: {checked_sv} / {total_sv} sinh viên.{warn_line}\n"
            f"Hệ thống sẽ ghi 8 cột điểm C - O - G - V (Part 1 & Part 2) vào file Excel đích:\n"
            f"• Tên file: {os.path.basename(excel_path)}\n"
            f"• Đường dẫn: {excel_path}\n"
            f"(Giữ nguyên 100% công thức tổng trong Excel, không ghi nhận xét AI).\n\n"
            f"Bạn có chắc chắn muốn Approve và xuất điểm sang file Excel này không?"
        )
        if not messagebox.askyesno("Xác nhận Approve & Xuất điểm sang Excel", confirm_msg):
            return

        while True:
            try:
                res = export_approved_scores_to_excel(
                    excel_path, self.review_items, include_comments=False, allow_fallback=False
                )
                saved_path = res["saved_path"]
                break
            except PermissionError:
                retry = messagebox.askretrycancel(
                    "File Excel đang mở trên máy",
                    f"Không thể ghi trực tiếp vì file Excel đang được mở trong Microsoft Excel:\n\n"
                    f"{excel_path}\n\n"
                    f"Vui lòng ĐÓNG cửa sổ Excel của file này rồi nhấn 'Retry' (Thử lại).\n"
                    f"(Hoặc nhấn 'Cancel' nếu muốn hủy thao tác).",
                )
                if not retry:
                    return
            except Exception as e:
                messagebox.showerror("Lỗi khi xuất điểm ra Excel", f"Chi tiết lỗi:\n{e}")
                return

        self.excel1_path_var.set(saved_path)
        self.excel2_path_var.set(saved_path)
        save_config({"last_excel_path": saved_path})
        self._update_target_excel_label()
        self._inspect_excel_tab2()

        open_now = messagebox.askyesno(
            "Đã lưu điểm ra Excel thành công",
            f"Đã ghi thành công 8 cột điểm C - O - G - V của {res['exported_count']} sinh viên vào file Excel:\n"
            f"{saved_path}\n\n"
            f"Bạn có muốn mở file Excel vừa lưu để kiểm tra ngay bây giờ không?",
        )
        if open_now:
            try:
                os.startfile(saved_path)
            except Exception:
                pass

    # ==================== XỬ LÝ TAB 1: CHẠY AI TẠO BẢN DUYỆT (ĐA LỚP CHẠY NGẦM ĐỘC LẬP) ====================
    def _start_ai_grading(self, only_ungraded: bool = False, target_code: str | None = None):
        api_key = self.api_key_var.get().strip()
        model_name = self.model_var.get().strip() or DEFAULT_MODEL
        self.rubric_content = load_marking_rules(self.rule_profile)

        if target_code is None:
            self._sync_active_class_tab(auto_rename_by_code=True)
            cls_code = self.active_class_code or extract_class_code(
                self.json_path_var.get().strip(), self.excel1_path_var.get().strip()
            )
            json_path, excel_path = resolve_or_recover_class_files(
                cls_code,
                self.json_path_var.get().strip(),
                self.excel1_path_var.get().strip(),
            )
        else:
            cls_code = target_code
            cls_entry = next((c for c in self.open_classes if c["code"] == cls_code), {})
            sess_existing = self._get_or_create_class_session(cls_code)
            json_path, excel_path = resolve_or_recover_class_files(
                cls_code,
                cls_entry.get("json_path", "") or sess_existing.get("json_path", ""),
                cls_entry.get("excel_path", "") or sess_existing.get("excel_path", ""),
            )

        if not api_key:
            messagebox.showwarning("Thiếu API Key", "Vui lòng nhập Gemini API Key ở Bước 1.")
            return
        if not json_path or not os.path.exists(json_path):
            messagebox.showwarning("Thiếu file JSON", f"Lớp {cls_code}: Vui lòng chọn file bai_lam_iuh.json hợp lệ.")
            return
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showwarning("Thiếu file Excel", f"Lớp {cls_code}: Vui lòng chọn file Excel Phiếu chấm (.xlsx) hợp lệ.")
            return

        try:
            validate_class_files_not_mixed(json_path, excel_path)
        except ValueError as mix_err:
            messagebox.showerror("Cảnh báo lệch Mã lớp giữa JSON và Excel", str(mix_err))
            return

        sess = self._get_or_create_class_session(cls_code)
        sess["json_path"] = json_path
        sess["excel_path"] = excel_path
        cls_item = next((c for c in self.open_classes if c["code"] == cls_code), None)
        if cls_item:
            cls_item["json_path"] = json_path
            cls_item["excel_path"] = excel_path

        if sess.get("ai_running"):
            return

        # Nếu đang có một lớp khác chạy AI, xếp lớp này vào hàng đợi tự động (Auto-Queue)
        if self.is_running_ai:
            if not any(c == cls_code for c, _ in self.ai_queue):
                self.ai_queue.append((cls_code, only_ungraded))
            sess["ai_queued"] = True
            sess["ai_status"] = "Đang chờ hàng đợi AI..."
            self._sync_ai_controls_for_active_class()
            self._render_class_bar()
            self._log1(f"[HÀNG ĐỢI AI] Đã xếp lớp {cls_code} vào hàng đợi. Sẽ tự động chấm ngay khi lớp hiện tại xong.")
            return

        arb_model = self.arbitrator_model_var.get().strip() or DEFAULT_ARBITRATOR_MODEL
        p2_enabled = bool(self.enable_pass2_var.get())

        save_config({
            "api_key": api_key,
            "model_name": model_name,
            "arbitrator_model": arb_model,
            "enable_pass2_arbitrator": p2_enabled,
            "last_json_path": json_path,
            "last_excel_path": excel_path,
            "rubric_text": self.rubric_content,
        })

        self.is_running_ai = True
        sess["ai_queued"] = False
        sess["ai_running"] = True
        sess["ai_done"] = 0
        sess["ai_status"] = "Đang chấm Hội đồng AI..."
        cls_stop_event: threading.Event = sess["stop_ai_event"]
        cls_stop_event.clear()

        self._show_pipeline_hud()
        init_hud_msg = (
            f"[{cls_code}] Khởi động Hội đồng AI Đa Mô hình: GK1={model_name} (High Thinking)"
            + (f"  ⇄  Trọng tài Vòng 2={arb_model}" if p2_enabled else "  (Chế độ 1 vòng tốc độ cao)")
        )
        self._reset_pipeline_hud_steps(init_hud_msg)
        self._sync_ai_controls_for_active_class()
        self._render_class_bar()
        self._log1(f"\n{init_hud_msg} -> Lưu độc lập vào phiên lớp {cls_code}...")

        def stage_cb(step_idx: int, state: str, badge_txt: str, detail_msg: str):
            def _ui_stage():
                self._update_pipeline_hud_step(step_idx, state, badge_txt, detail_msg)
                if detail_msg:
                    self._log1(f"[{cls_code}] {detail_msg}")
            self.root.after(0, _ui_stage)

        def progress_cb(completed, total, mssv, ten, kq: KetQuaCham | None, err):
            def _ui():
                sess["ai_done"] = completed
                sess["ai_total"] = total
                sess["ai_status"] = f"Đã chấm: {completed}/{total} SV"
                if self.active_class_code == cls_code:
                    self.ai_progress["maximum"] = max(total, 1)
                    self.ai_progress["value"] = completed
                    self.ai_status_var.set(sess["ai_status"])
                self._render_class_bar()
                if err is not None:
                    self._log1(f"[{cls_code} LỖI {completed}/{total}] MSSV {mssv} ({ten}): {err}")
                elif kq is not None:
                    f1 = (kq.p1_content + kq.p1_org + kq.p1_grammar + kq.p1_vocab) / 4.0
                    f2 = (kq.p2_content + kq.p2_org + kq.p2_grammar + kq.p2_vocab) / 4.0
                    self._log1(
                        f"[{cls_code} OK {completed}/{total}] {mssv} - {ten} | "
                        f"P1: ({kq.p1_content}-{kq.p1_org}-{kq.p1_grammar}-{kq.p1_vocab})={f1:g} | "
                        f"P2: ({kq.p2_content}-{kq.p2_org}-{kq.p2_grammar}-{kq.p2_vocab})={f2:g}"
                    )
            self.root.after(0, _ui)

        def worker():
            try:
                summary = chay_cham_ai_tao_ban_duyet(
                    api_key=api_key,
                    json_path=json_path,
                    excel_path=excel_path,
                    rubric_text=self.rubric_content,
                    progress_cb=progress_cb,
                    stop_event=cls_stop_event,
                    model_name=model_name,
                    only_ungraded=only_ungraded,
                    enable_pass2_arbitrator=p2_enabled,
                    arbitrator_model=arb_model,
                    stage_cb=stage_cb,
                )

                def _done():
                    self.is_running_ai = False
                    sess["ai_running"] = False
                    items_done = summary.get("items", [])
                    sess["review_items"] = items_done
                    if cls_item is not None:
                        cls_item["checked_count"] = sum(1 for it in items_done if it.get("checked"))
                        cls_item["total_count"] = len(items_done)

                    if cls_stop_event.is_set():
                        sess["ai_status"] = "Đã dừng theo yêu cầu."
                        self._log1(f"[{cls_code}] Đã dừng chấm. Đã lưu tạm {summary['graded_count']} SV vào phiên kiểm duyệt của lớp {cls_code}.")
                    else:
                        sess["ai_status"] = f"Hoàn tất {summary['graded_count']}/{summary['total_matched']} SV."
                        self._log1(
                            f"\n[{cls_code}] HOÀN TẤT CHẤM AI! Đã lưu độc lập {len(items_done)} sinh viên vào phiên lớp {cls_code}."
                        )

                    # Chỉ cập nhật màn hình nếu giảng viên đang mở chính lớp đó; nếu đang kiểm lớp khác thì giữ nguyên màn hình!
                    if self.active_class_code == cls_code:
                        self.review_items = items_done
                        self._update_match_summary_from_items(items_done)
                        self._apply_review_filter()

                    self._sync_ai_controls_for_active_class()
                    self._persist_class_bar_state()
                    self._render_class_bar()
                    self._run_next_queued_ai()
                self.root.after(0, _done)
            except Exception as e:
                def _err():
                    self.is_running_ai = False
                    sess["ai_running"] = False
                    sess["ai_status"] = "Gặp lỗi."
                    self._sync_ai_controls_for_active_class()
                    self._render_class_bar()
                    self._log1(f"[{cls_code}] Lỗi khi chấm AI: {e}")
                    messagebox.showerror(f"Lỗi khi chấm AI ({cls_code})", f"Chi tiết lỗi:\n{e}")
                    self._run_next_queued_ai()
                self.root.after(0, _err)

        threading.Thread(target=worker, daemon=True).start()

    def _run_next_queued_ai(self):
        while self.ai_queue and not self.is_running_ai:
            next_code, next_ungraded = self.ai_queue.pop(0)
            sess = self.class_sessions.get(next_code)
            if sess:
                sess["ai_queued"] = False
            if any(c["code"] == next_code for c in self.open_classes):
                self._start_ai_grading(only_ungraded=next_ungraded, target_code=next_code)
                break

    def _stop_ai_grading(self):
        sess = self.class_sessions.get(self.active_class_code) if self.active_class_code else None
        if sess and sess.get("ai_queued"):
            sess["ai_queued"] = False
            sess["ai_status"] = "Đã hủy hàng đợi."
            self.ai_queue = [(c, u) for (c, u) in self.ai_queue if c != self.active_class_code]
            self._sync_ai_controls_for_active_class()
            self._render_class_bar()
            return
        if sess and sess.get("ai_running"):
            sess["stop_ai_event"].set()
            sess["ai_status"] = "Đang dừng các luồng..."
            self.ai_status_var.set("Đang dừng các luồng...")
            return
        # Nếu đang đứng ở tab khác mà bấm dừng, dừng lớp đang chạy AI
        for c_code, s_obj in self.class_sessions.items():
            if s_obj.get("ai_running"):
                s_obj["stop_ai_event"].set()
                s_obj["ai_status"] = "Đang dừng các luồng..."
                self._log1(f"[{c_code}] Đang dừng tiến trình chấm AI...")
                break

    # ==================== XỬ LÝ TAB 3: TỰ ĐỘNG NHẬP ĐIỂM LÊN MOODLE ====================
    def _inspect_excel_tab2(self):
        path = self.excel2_path_var.get().strip()
        if not path or not os.path.exists(path):
            self.excel_summary_var.set("Chưa tìm thấy file Excel. Vui lòng chọn file hợp lệ.")
            return
        try:
            cols, students = detect_columns_and_parse(path)
            for child in self.web_tree.get_children():
                self.web_tree.delete(child)

            if not students:
                self.excel_summary_var.set("Không tìm thấy dòng sinh viên nào có MSSV hợp lệ trong file Excel này.")
                return

            valid_web_mssv = None
            json_path = self.json_path_var.get().strip()
            if json_path and os.path.exists(json_path):
                try:
                    with open(json_path, "r", encoding="utf-8") as f:
                        ds_json = json.load(f)
                    valid_web_mssv = {normalize_mssv(x.get("mssv")) for x in ds_json if normalize_mssv(x.get("mssv"))}
                except Exception:
                    valid_web_mssv = None

            for st in students:
                mssv = st["mssv"]
                is_absent = valid_web_mssv is not None and mssv not in valid_web_mssv
                status_txt = "Bỏ qua (Vắng thi trên Moodle)" if is_absent else "Sẵn sàng nhập"
                self.web_tree.insert(
                    "",
                    tk.END,
                    iid=mssv,
                    values=(st["row"], mssv, st["ten"], st["q17"], st["q18"], status_txt),
                )

            st0 = students[0]
            msg = (
                f"Đã nhận diện: {len(students)} sinh viên (từ dòng {st0['row']})  |  "
                f"MSSV: {col_label(cols['mssv'])}  |  Họ tên: {col_label(cols['ten'])}  |  "
                f"FINAL SCORE 1 (Q17): {col_label(cols['q17'])}  |  FINAL SCORE 2 (Q18): {col_label(cols['q18'])}"
            )
            self.excel_summary_var.set(msg)
        except Exception as e:
            self.excel_summary_var.set(f"Lỗi khi đọc file Excel: {e}")

    def _on_web_tree_select(self, _event):
        sel = self.web_tree.selection()
        if not sel:
            return
        vals = self.web_tree.item(sel[0], "values")
        if vals and len(vals) >= 2:
            self.start_filter_var.set(str(vals[1]))

    def _open_excel_for_review(self):
        path = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Chưa có file Excel", "Vui lòng chọn file Excel trước khi mở.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Không thể mở file", f"Lỗi khi mở file Excel:\n{e}")

    def _open_web_log_file(self):
        excel_path = self.excel2_path_var.get().strip() or self.excel1_path_var.get().strip()
        log_dir = os.path.dirname(excel_path) if excel_path else get_app_dir()
        log_path = os.path.join(log_dir, "nhat_ky_cham_diem.txt")
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            messagebox.showinfo("Chưa có nhật ký", f"Chưa có file nhật ký tại:\n{log_path}")

    def _reload_scores_from_excel(self):
        """Nạp lại điểm mới nhất trực tiếp từ file Excel (.xlsx) khi giảng viên vừa sửa điểm trong Excel để khớp với GV2."""
        raw_json = self.json_path_var.get().strip()
        raw_excel = self.excel1_path_var.get().strip() or self.excel2_path_var.get().strip()
        json_p, excel_p = resolve_or_recover_class_files(self.active_class_code, raw_json, raw_excel)
        if not excel_p or not os.path.exists(excel_p):
            messagebox.showwarning("Chưa có file Excel", "Vui lòng chọn file Excel Phiếu chấm (.xlsx) trước khi nạp lại điểm.")
            return
        try:
            items = build_review_items_from_excel_and_json(
                excel_p,
                json_p if (json_p and os.path.exists(json_p)) else "",
                class_code=self.active_class_code,
                prefer_excel_scores=True,
            )
            self.review_items = items
            if self.active_class_code:
                sess = self._get_or_create_class_session(self.active_class_code)
                sess["review_items"] = items
            self._auto_save_review_session()
            self._update_match_summary_from_items(items)
            self._apply_review_filter()
            self._inspect_excel_tab2()
            messagebox.showinfo(
                "Đã đồng bộ từ Excel",
                f"Đã nạp lại điểm 8 cột C-O-G-V và điểm quy đổi Q17, Q18 mới nhất từ file:\n{os.path.basename(excel_p)}",
            )
        except Exception as e:
            messagebox.showerror("Lỗi nạp điểm từ Excel", f"Không thể đồng bộ từ file Excel:\n{e}")

    def _collect_multi_class_grades(self, only_active_class: bool = True) -> tuple[dict[str, dict], list[str]]:
        """
        Thu thập điểm Q17 & Q18 (theo MSSV) để đóng gói mã JS nhập điểm Moodle:
        - Mặc định `only_active_class=True`: Chỉ đóng gói đúng lớp đang chọn (`self.active_class_code`), không gộp tràn sang các lớp khác đang mở.
        - Tự động nhận diện nếu file Excel vừa được sửa mới hơn file _review.json.
        """
        grades_map: dict[str, dict] = {}
        loaded_classes: list[str] = []

        class_entries = list(self.open_classes)
        cur_json = self.json_path_var.get().strip()
        cur_excel = self.excel2_path_var.get().strip() or self.excel1_path_var.get().strip()
        if not class_entries and (cur_json or cur_excel):
            class_entries.append({
                "code": self.active_class_code or extract_class_code(cur_json, cur_excel),
                "json_path": cur_json,
                "excel_path": cur_excel,
            })

        if only_active_class and self.active_class_code:
            active_entries = [
                e for e in class_entries
                if (e.get("code") or extract_class_code(e.get("json_path", ""), e.get("excel_path", ""))) == self.active_class_code
            ]
            if not active_entries and (cur_json or cur_excel):
                active_entries = [{
                    "code": self.active_class_code,
                    "json_path": cur_json,
                    "excel_path": cur_excel,
                }]
            class_entries = active_entries

        start_input = self.start_filter_var.get().strip()

        for entry in class_entries:
            code = entry.get("code") or extract_class_code(entry.get("json_path", ""), entry.get("excel_path", ""))
            j_path, e_path = resolve_or_recover_class_files(
                code,
                (entry.get("json_path") or "").strip(),
                (entry.get("excel_path") or "").strip(),
            )
            is_current_active = code == self.active_class_code
            added_for_class = 0

            # Kiểm tra xem giảng viên có vừa mở file Excel sửa điểm (để khớp với GV2) mới hơn file _review.json không
            rev_path = get_review_session_path(j_path, e_path)
            excel_newer_on_disk = False
            if e_path and os.path.exists(e_path):
                if not rev_path or not os.path.exists(rev_path):
                    excel_newer_on_disk = True
                else:
                    try:
                        if os.path.getmtime(e_path) > os.path.getmtime(rev_path) + 0.1:
                            excel_newer_on_disk = True
                    except OSError:
                        pass

            if excel_newer_on_disk and e_path and os.path.exists(e_path):
                try:
                    synced_items = build_review_items_from_excel_and_json(
                        e_path,
                        j_path if (j_path and os.path.exists(j_path)) else "",
                        class_code=code,
                        prefer_excel_scores=True,
                    )
                    if synced_items:
                        if is_current_active:
                            self.review_items = synced_items
                        sess_m = self._get_or_create_class_session(code)
                        sess_m["review_items"] = synced_items
                        save_review_session(j_path, e_path, synced_items, class_code=code)
                except Exception:
                    pass
            elif is_current_active:
                self._auto_save_review_session()

            # 1. Đọc từ file Excel nếu có
            if e_path and os.path.exists(e_path):
                try:
                    _, students = detect_columns_and_parse(e_path)
                    if is_current_active and start_input and start_input != "10":
                        students = filter_start_students(students, start_input)
                    valid_web_mssv = None
                    if j_path and os.path.exists(j_path):
                        try:
                            with open(j_path, "r", encoding="utf-8") as f:
                                ds_json = json.load(f)
                            valid_web_mssv = {
                                normalize_mssv(x.get("mssv")) for x in ds_json if normalize_mssv(x.get("mssv"))
                            }
                        except Exception:
                            valid_web_mssv = None
                    for st in students:
                        m = normalize_mssv(st.get("mssv"))
                        if not m:
                            continue
                        if valid_web_mssv is not None and m not in valid_web_mssv:
                            continue
                        grades_map[m] = {
                            "ten": st.get("ten", ""),
                            "q17": str(st.get("q17", "0")),
                            "q18": str(st.get("q18", "0")),
                            "class_code": code,
                        }
                        added_for_class += 1
                except Exception:
                    pass

            # 2. Cập nhật điểm từ bộ nhớ phiên (self.class_sessions) hoặc file *_review.json của lớp đó (nếu Excel không mới hơn)
            if not excel_newer_on_disk:
                rev_items = None
                sess_mem = self.class_sessions.get(code)
                if is_current_active and self.review_items:
                    rev_items = self.review_items
                elif sess_mem and sess_mem.get("review_items"):
                    rev_items = sess_mem["review_items"]
                else:
                    loaded_rev = load_review_session(j_path, e_path, class_code=code)
                    if loaded_rev and isinstance(loaded_rev.get("items"), list):
                        rev_items = loaded_rev["items"]

                if rev_items:
                    filtered_rev = rev_items
                    if is_current_active and start_input and start_input != "10":
                        mssv_cand = normalize_mssv(start_input)
                        if mssv_cand:
                            for idx_r, it_r in enumerate(rev_items):
                                if normalize_mssv(it_r.get("mssv")) == mssv_cand:
                                    filtered_rev = rev_items[idx_r:]
                                    break
                        elif start_input.isdigit():
                            min_r = int(start_input)
                            filtered_rev = [it_r for it_r in rev_items if int(it_r.get("row") or 0) >= min_r]

                    for it in filtered_rev:
                        if it.get("absent"):
                            continue
                        m = normalize_mssv(it.get("mssv"))
                        if not m:
                            continue
                        f1, f2, _ = self._calc_item_scores(it)
                        grades_map[m] = {
                            "ten": it.get("ten", ""),
                            "q17": format_score(f1),
                            "q18": format_score(f2),
                            "class_code": code,
                        }
                        added_for_class += 1

            if added_for_class > 0 and code not in loaded_classes:
                loaded_classes.append(code)

        return grades_map, loaded_classes

    def _copy_js_paced_grade_submitter(self):
        grades_map, loaded_classes = self._collect_multi_class_grades(only_active_class=True)
        if not grades_map:
            messagebox.showwarning(
                "Chưa có dữ liệu điểm",
                "Chưa tìm thấy điểm sinh viên nào của lớp đang chọn.\n"
                "Vui lòng chọn file Excel/JSON hoặc kiểm duyệt điểm trước.",
            )
            return

        speed_mult = float(self.speed_var.get() or 1.0)
        min_sec = round(3.0 * speed_mult, 1)
        max_sec = round(5.0 * speed_mult, 1)
        js_code = build_js_paced_grade_submitter(
            grades_map,
            min_delay_sec=min_sec,
            max_delay_sec=max_sec,
            class_codes=loaded_classes,
        )
        pyperclip.copy(js_code)

        classes_str = ", ".join(loaded_classes) if loaded_classes else (self.active_class_code or "Lớp hiện tại")
        self._log2("------------------------------------------------------------------")
        self._log2(
            f"Đã copy Mã JS Nhập Điểm Chạy Ngầm (F12) cho {len(grades_map)} sinh viên "
            f"lớp {classes_str} | Nhịp tự nhiên: {min_sec}s – {max_sec}s/câu."
        )
        messagebox.showinfo(
            "Đã Copy Mã JS Nhập Điểm Chạy Ngầm (F12)",
            f"Đã đóng gói điểm của {len(grades_map)} sinh viên (Lớp đang chọn: {classes_str}) vào Clipboard!\n\n"
            f"Cách chạy ngầm (An toàn 100%, không chiếm chuột/bàn phím):\n"
            f"1. Mở trang Bảng điểm của lớp {classes_str} trên luyenthinn.iuh.edu.vn.\n"
            f"2. Nhấn F12 -> chuyển sang thẻ Console.\n"
            f"3. Dán (Ctrl + V) và nhấn Enter.\n"
            f"4. Trình duyệt sẽ tự động đối chiếu đúng MSSV của lớp {classes_str} và lưu điểm từng câu với nhịp nghỉ tự nhiên {min_sec}–{max_sec}s/câu.\n"
            f"   Bạn có thể thu nhỏ trình duyệt và làm việc khác bình thường!",
        )

    def _start_web_upload(self):
        if self.is_running_web:
            return
        excel_path = self.excel2_path_var.get().strip()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showwarning("Thiếu file Excel", "Vui lòng chọn file Excel bảng điểm đã Approve.")
            return

        try:
            _, all_students = detect_columns_and_parse(excel_path)
        except Exception as e:
            messagebox.showerror("Lỗi đọc Excel", f"Không thể đọc file Excel:\n{e}")
            return

        json_path = self.json_path_var.get().strip()
        valid_web_mssv = None
        if json_path and os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    ds_json = json.load(f)
                valid_web_mssv = {normalize_mssv(x.get("mssv")) for x in ds_json if normalize_mssv(x.get("mssv"))}
            except Exception:
                valid_web_mssv = None

        start_input = self.start_filter_var.get().strip()
        students = filter_start_students(all_students, start_input)
        if not students:
            messagebox.showwarning("Danh sách trống", "Không có sinh viên nào trong phạm vi bắt đầu đã chọn.")
            return

        speed_mult = float(self.speed_var.get() or 1.0)
        self.is_running_web = True
        self.stop_web_event.clear()
        self.btn_run_web.config(state=tk.DISABLED)
        self.btn_stop_web.config(state=tk.NORMAL)
        self.web_progress["maximum"] = len(students)
        self.web_progress["value"] = 0

        log_path = os.path.join(os.path.dirname(excel_path), "nhat_ky_cham_diem.txt")
        self._log2("------------------------------------------------------------------")
        self._log2(
            f"Chuẩn bị nhập điểm cho {len(students)} sinh viên "
            f"(Bắt đầu từ dòng {students[0]['row']}: {students[0]['mssv']} - {students[0]['ten']} | Tốc độ: {speed_mult}x)"
        )
        self._log2("Bạn có 5 GIÂY: Hãy CLICK CHUỘT 1 CÁI vào khoảng trắng trên trang Bảng điểm của trình duyệt.")

        def _update_web_row_status(mssv_key: str, status_str: str):
            if self.web_tree.exists(mssv_key):
                vals = list(self.web_tree.item(mssv_key, "values"))
                if len(vals) >= 6:
                    vals[5] = status_str
                    self.web_tree.item(mssv_key, values=vals)
                    self.web_tree.see(mssv_key)

        def worker():
            for sec in range(5, 0, -1):
                if self.stop_web_event.is_set():
                    break
                self.root.after(0, lambda s=sec: self.web_status_var.set(f"Đếm ngược: {s} giây (Click vào trình duyệt)"))
                time.sleep(1.0)

            if self.stop_web_event.is_set():
                self.root.after(0, self._finish_web_ui, "Đã hủy trước khi chạy.")
                return

            current_sv_info = ""
            log_file = open(log_path, "a", encoding="utf-8")
            try:
                for idx, st in enumerate(students, 1):
                    if self.stop_web_event.is_set():
                        break
                    mssv, ten, row_num = st["mssv"], st["ten"], st["row"]
                    if valid_web_mssv is not None and mssv not in valid_web_mssv:
                        self.root.after(
                            0,
                            lambda i=idx, r=row_num, m=mssv, t=ten: (
                                self._log2(f"[Bỏ qua {i}/{len(students)}] (Dòng {r}) {m} ({t}) - Vắng thi trên Moodle."),
                                _update_web_row_status(m, "Đã bỏ qua (Vắng thi)"),
                            ),
                        )
                        self.root.after(0, lambda i=idx: self.web_progress.configure(value=i))
                        continue

                    for phan, diem in [("Q17 (Part 1)", st["q17"]), ("Q18 (Part 2)", st["q18"])]:
                        if self.stop_web_event.is_set():
                            break
                        current_sv_info = f"Dòng {row_num} | {mssv} ({ten}) - {phan}"
                        self.root.after(
                            0,
                            lambda i=idx, r=row_num, m=mssv, t=ten, p=phan, d=diem: (
                                self.web_status_var.set(f"Đang nhập [{i}/{len(students)}]: {m} - {p}"),
                                self._log2(f"[{i}/{len(students)}] (Dòng {r}) Đang nhập {m} ({t}) - {p}: điểm = {d}..."),
                                _update_web_row_status(m, f"Đang nhập {p}..."),
                            ),
                        )
                        cham_mot_cau_web(mssv, diem, speed_mult=speed_mult)
                        log_file.write(
                            f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | "
                            f"Dòng {row_num} | {mssv} | {ten} | {phan} = {diem}\n"
                        )
                        log_file.flush()

                    self.root.after(
                        0,
                        lambda i=idx, m=mssv: (
                            self.web_progress.configure(value=i),
                            _update_web_row_status(m, "Đã nhập xong"),
                        ),
                    )

                if self.stop_web_event.is_set():
                    self.root.after(
                        0,
                        lambda: (
                            self._log2(f"Đã dừng theo yêu cầu tại: {current_sv_info}"),
                            self._finish_web_ui("Đã dừng."),
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda: (
                            self._log2(f"ĐÃ HOÀN TẤT NHẬP ĐIỂM TOÀN BỘ DANH SÁCH. Nhật ký lưu tại: {log_path}"),
                            self._finish_web_ui("Hoàn tất 100%."),
                            messagebox.showinfo("Hoàn tất", f"Đã nhập điểm xong cho danh sách sinh viên.\nNhật ký lưu tại:\n{log_path}"),
                        ),
                    )
            except pyautogui.FailSafeException:
                self.root.after(
                    0,
                    lambda: (
                        self._log2(
                            f"\nĐÃ DỪNG KHẨN CẤP DO VẨY CHUỘT VÀO GÓC MÀN HÌNH.\n"
                            f"  - Vị trí đang dừng: {current_sv_info}\n"
                            f"  - Nhật ký đã lưu tại: {log_path}"
                        ),
                        self._finish_web_ui("Dừng khẩn cấp (FailSafe)."),
                        messagebox.showwarning(
                            "Đã Dừng Khẩn Cấp",
                            f"Đã dừng khẩn cấp do bạn vẩy chuột vào góc màn hình.\n\n"
                            f"Vị trí đang dừng:\n{current_sv_info}",
                        ),
                    ),
                )
            except Exception as e:
                self.root.after(
                    0,
                    lambda err=e: (
                        self._log2(f"Có lỗi xảy ra khi đang nhập ({current_sv_info}): {err}"),
                        self._finish_web_ui("Gặp lỗi."),
                        messagebox.showerror("Lỗi khi nhập điểm", f"Vị trí: {current_sv_info}\nChi tiết lỗi: {err}"),
                    ),
                )
            finally:
                log_file.close()

        threading.Thread(target=worker, daemon=True).start()

    def _finish_web_ui(self, status_text: str):
        self.is_running_web = False
        self.btn_run_web.config(state=tk.NORMAL)
        self.btn_stop_web.config(state=tk.DISABLED)
        self.web_status_var.set(status_text)

    def _stop_web_upload(self):
        if self.is_running_web:
            self.stop_web_event.set()
            self.web_status_var.set("Đang dừng sau câu hiện tại...")


def main():
    root = tk.Tk()
    app = IUHWritingAISuiteApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

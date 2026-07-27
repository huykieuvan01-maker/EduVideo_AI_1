import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const SYSTEM_INSTRUCTION = `
Bạn là **EduVideo_AI**, một Chuyên gia Kiến trúc Giáo dục và Sư phạm video tích hợp. Bạn đóng vai trò là "bộ não" trung tâm điều phối toàn bộ quy trình thiết kế kịch bản bài giảng video, tích hợp quiz tương tác, phân tích tri thức và chẩn đoán tiến độ học tập của học sinh thông qua video học tập.

Mục tiêu cốt lõi của bạn là hỗ trợ giáo viên thiết kế video bài giảng có chiều sâu, giúp học sinh học chủ động qua video, và cung cấp cho phụ huynh các phân tích tiến độ trực quan.

Bạn sẽ được gọi qua 3 bước xử lý (Step 1, Step 2, Step 3) dựa trên vai trò của người dùng (Teacher, Student, Parent). Hãy tuân thủ nghiêm ngặt các hướng dẫn và định dạng đầu ra cho từng bước như sau:

---
VAI TRÒ: TEACHER (Giáo viên / Nhà sáng tạo nội dung)
- **Step 1: Video Outline & Script (Kịch bản Video Chi tiết)**
  Yêu cầu: Lập dàn ý bài học và kịch bản video chi tiết chia theo từng cảnh. Với mỗi cảnh, chỉ rõ:
  - [Thời gian dự kiến]
  - [Chỉ dẫn bảng trắng / Hình ảnh hiển thị trên video]
  - [Lời thoại chi tiết cho Giáo viên ảo / Avatar]
- **Step 2: In-Video Quiz Overlays & Timestamps (Câu hỏi tương tác trong Video)**
  Yêu cầu: Tạo ra 3-4 câu hỏi trắc nghiệm (Quiz) xen kẽ vào các mốc thời gian cụ thể của video để học sinh trả lời khi đang xem bài giảng. Mỗi câu hỏi gồm:
  - Mốc thời gian xuất hiện (ví dụ: [02:15], [05:40])
  - Nội dung câu hỏi và các đáp án lựa chọn (A, B, C, D)
  - Đáp án đúng và một Gợi ý thông minh (Smart Hint) giúp định hướng tư duy chứ không trực tiếp đưa ra câu trả lời.
- **Step 3: Supplemental Materials & Exercise Matrix (Tài liệu bổ trợ & Ma trận bài tập)**
  Yêu cầu: Tạo ma trận bài tập luyện tập bổ sung phân cấp theo 3 cấp độ: Nhận biết, Thông hiểu, Vận dụng. Cung cấp thêm 1 lời khuyên sư phạm ngắn gọn để tối ưu hóa hiệu quả bài học.

---
VAI TRÒ: STUDENT (Học sinh học tập chủ động)
- **Step 1: Key Concept Map & Summary (Tóm tắt & Sơ đồ khái niệm)**
  Yêu cầu: Tóm tắt các khái niệm cốt lõi của bài học / video dưới dạng sơ đồ khái niệm trực quan (dùng ký tự văn bản / ASCII hoặc dạng danh sách có phân tầng liên kết) và giải thích ngắn gọn ý nghĩa từng khái niệm.
- **Step 2: Interactive Notes & Chapters (Ghi chú chi tiết & Đề mục bài học)**
  Yêu cầu: Tạo ghi chú bài học có phân chia chương/mục kèm mốc thời gian gợi ý. Nêu bật các công thức, định lý hoặc quy tắc quan trọng cần ghi nhớ.
- **Step 3: Active Recall & Flashcard Questions (Học chủ động & Câu hỏi phản xạ)**
  Yêu cầu: Tạo bộ câu hỏi phản xạ (Active Recall) để tự ôn tập, kèm theo 3 cấp độ gợi ý (Smart Hint) cho mỗi câu hỏi nếu học sinh chưa tìm ra lời giải.

---
VAI TRÒ: PARENT (Phụ huynh theo dõi & Đồng hành)
- **Step 1: Learning Progress Pulse (Nhịp độ học tập & Tiến trình)**
  Yêu cầu: Mô phỏng nhịp độ học tập của học sinh qua video (Tỷ lệ xem hết, mức độ tập trung ở các phần khó của video) dưới dạng bảng thống kê hoặc biểu đồ ký tự đơn giản.
- **Step 2: Gap Analysis (Chẩn đoán lỗ hổng kiến thức)**
  Yêu cầu: Phân tích các lỗ hổng kiến thức tiềm ẩn của học sinh khi học chủ đề này qua video (dựa trên các lỗi sai thường gặp khi trả lời quiz).
- **Step 3: Personalized Remedial Path (Lộ trình cải thiện cá nhân hóa)**
  Yêu cầu: Đề xuất lộ trình ôn tập và luyện tập 7 ngày tiếp theo cụ thể cho con, kèm theo gợi ý các phân đoạn video con cần xem lại để lấp lỗ hổng.

Hãy trả lời bằng Tiếng Việt, phong cách chuyên nghiệp, sư phạm và khuyến khích học tập. Tránh dài dòng, tập trung vào cấu trúc trực quan dễ đọc.
`;

const app = express();
app.use(express.json());

const PORT = 3000;

let aiClient: GoogleGenAI | null = null;
function getAIClient(): GoogleGenAI {
  if (!aiClient) {
    if (!process.env.GEMINI_API_KEY) {
      throw new Error("GEMINI_API_KEY is required");
    }
    aiClient = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  }
  return aiClient;
}

app.post("/api/educore/generate", async (req, res) => {
  try {
    const { prompt, mode, step, model } = req.body;
    
    // Accept key from headers (x-api-key) or query or request body
    const apiKey = req.headers["x-api-key"] as string || req.body.apiKey || process.env.GEMINI_API_KEY;

    if (!apiKey) {
      return res.status(401).json({ error: "Yêu cầu GEMINI_API_KEY. Vui lòng thiết lập trong Cài đặt." });
    }

    if (!prompt) {
      return res.status(400).json({ error: "Prompt is required" });
    }

    const ai = new GoogleGenAI({ apiKey });
    const targetModel = model || "gemini-2.5-flash";
    
    const contextPrompt = `
Vai trò người dùng hiện tại đang là: ${mode} (Teacher, Student, Parent).
Đang thực hiện bước xử lý: BƯỚC ${step} (Step ${step}).

Dựa trên tài liệu/yêu cầu sau đây, hãy thực hiện nhiệm vụ của EduVideo_AI cho bước này:
${prompt}

---
YÊU CẦU ĐẶC BIỆT VỀ DỮ LIỆU ĐẦU RA (Quan trọng cho Frontend):
Hãy chèn thêm một khối JSON ở cuối câu trả lời của bạn, nằm trong thẻ code: \`\`\`json và \`\`\`. Khối JSON này chỉ chứa dữ liệu cấu trúc cho bước này như sau:

- Nếu bạn ở vai trò Teacher và Step 1 (Video Script):
  Khối JSON có dạng:
  {
    "slideData": [
      { "title": "Tên slide", "bulletPoints": ["Ý chính 1", "Ý chính 2"], "notes": "Lời thoại giáo viên" }
    ]
  }

- Nếu bạn ở vai trò Teacher và Step 3 (Supplemental Materials):
  Khối JSON có dạng:
  {
    "worksheetData": {
      "title": "Tên tài liệu",
      "exercises": [
        { "question": "Câu hỏi tự luận/trắc nghiệm", "options": ["A...", "B...", "C...", "D..."], "answer": "Đáp án đúng / Gợi ý" }
      ]
    }
  }

- Nếu bạn ở vai trò Student và Step 1 (Concept Map):
  Khối JSON có dạng:
  {
    "conceptMapData": {
      "nodes": [
        { "id": "1", "label": "Khái niệm chính", "type": "main" },
        { "id": "2", "label": "Khái niệm phụ", "type": "sub" }
      ],
      "links": [
        { "source": "1", "target": "2" }
      ]
    }
  }

- Nếu bạn ở vai trò Student và Step 3 (Active Recall):
  Khối JSON có dạng:
  {
    "quizData": [
      { "question": "Câu hỏi ôn luyện", "hints": ["Gợi ý cấp 1", "Gợi ý cấp 2"], "answer": "Đáp án ẩn" }
    ]
  }

- Nếu bạn ở vai trò Parent và Step 1 (Learning Pulse):
  Khối JSON có dạng:
  {
    "pulseData": [
      { "day": "Thứ 2", "watchPercent": 45, "score": 70 },
      { "day": "Thứ 3", "watchPercent": 60, "score": 75 },
      { "day": "Thứ 4", "watchPercent": 80, "score": 80 },
      { "day": "Thứ 5", "watchPercent": 85, "score": 85 },
      { "day": "Thứ 6", "watchPercent": 90, "score": 85 },
      { "day": "Thứ 7", "watchPercent": 95, "score": 90 },
      { "day": "CN", "watchPercent": 100, "score": 92 }
    ]
  }

- Nếu bạn ở vai trò Parent và Step 2 (Gap Analysis):
  Khối JSON có dạng:
  {
    "gapData": [
      { "subject": "Khái niệm cơ bản", "score": 85 },
      { "subject": "Áp dụng công thức", "score": 50 },
      { "subject": "Tư duy nâng cao", "score": 30 },
      { "subject": "Giải quyết vấn đề", "score": 60 },
      { "subject": "Thuyết trình/Giao tiếp", "score": 75 }
    ]
  }

Nếu không thuộc các bước trên, hãy chèn khối JSON rỗng: {}
`;

    const response = await ai.models.generateContent({
      model: targetModel,
      contents: contextPrompt,
      config: {
        systemInstruction: SYSTEM_INSTRUCTION,
        temperature: 0.7,
      },
    });

    res.json({ result: response.text });
  } catch (error: any) {
    console.error("AI Error:", error);
    res.status(500).json({ error: error.message || "Something went wrong" });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();

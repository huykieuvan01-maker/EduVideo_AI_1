import { useState, useEffect } from 'react';
import Markdown from 'react-markdown';
import pptxgen from 'pptxgenjs';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import { 
  Video, 
  Sparkles, 
  Loader2, 
  Send, 
  Settings, 
  Key, 
  ExternalLink, 
  CheckCircle2, 
  XCircle, 
  AlertCircle, 
  BookOpen, 
  GraduationCap, 
  Users, 
  Play, 
  FileText, 
  HelpCircle, 
  ArrowRight,
  Download,
  Flame,
  Award
} from 'lucide-react';

type Role = 'Teacher' | 'Student' | 'Parent';

interface StepState {
  status: 'idle' | 'waiting' | 'loading' | 'success' | 'error';
  model: string;
  result: string;
  error?: string;
  metadata?: any; // Parsed JSON metadata
}

const modelsChain = [
  'gemini-3-flash-preview',
  'gemini-3-pro-preview',
  'gemini-2.5-flash'
];

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

export default function App() {
  const [activeTab, setActiveTab] = useState<Role>('Teacher');
  const [inputData, setInputData] = useState('');
  const [loading, setLoading] = useState(false);
  
  // Settings State
  const [apiKey, setApiKey] = useState('');
  const [selectedModel, setSelectedModel] = useState('gemini-3-flash-preview');
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [tempKey, setTempKey] = useState('');
  const [tempModel, setTempModel] = useState('gemini-3-flash-preview');
  
  // Pipeline Steps State
  const [step1, setStep1] = useState<StepState>({ status: 'idle', model: '', result: '' });
  const [step2, setStep2] = useState<StepState>({ status: 'idle', model: '', result: '' });
  const [step3, setStep3] = useState<StepState>({ status: 'idle', model: '', result: '' });
  const [globalError, setGlobalError] = useState<string | null>(null);

  // Active Recall Game Toggle
  const [playGame, setPlayGame] = useState(false);

  // Load configuration from localStorage
  useEffect(() => {
    const savedKey = localStorage.getItem('gemini_api_key') || '';
    const savedModel = localStorage.getItem('gemini_selected_model') || 'gemini-3-flash-preview';
    setApiKey(savedKey);
    setSelectedModel(savedModel);
    setTempKey(savedKey);
    setTempModel(savedModel);
  }, []);

  const saveSettings = (keyToSave: string, modelToSave: string) => {
    localStorage.setItem('gemini_api_key', keyToSave);
    localStorage.setItem('gemini_selected_model', modelToSave);
    setApiKey(keyToSave);
    setSelectedModel(modelToSave);
    setIsSettingsOpen(false);
  };

  // Helper to extract JSON from AI response
  const parseJSONMetadata = (text: string) => {
    try {
      const match = text.match(/```json\s*([\s\S]*?)\s*```/);
      if (match && match[1]) {
        return JSON.parse(match[1].trim());
      }
    } catch (e) {
      console.error("Failed to parse metadata JSON", e);
    }
    return null;
  };

  // Strip JSON metadata block for clean markdown rendering
  const cleanMarkdownText = (text: string) => {
    return text.replace(/```json\s*[\s\S]*?\s*```/g, '').trim();
  };

  const handleGenerate = async () => {
    if (!inputData.trim()) return;
    if (!apiKey.trim()) {
      setIsSettingsOpen(true);
      return;
    }

    setLoading(true);
    setGlobalError(null);
    setPlayGame(false);

    // Reset steps
    setStep1({ status: 'loading', model: selectedModel, result: '' });
    setStep2({ status: 'waiting', model: '', result: '' });
    setStep3({ status: 'waiting', model: '', result: '' });

    // Step 1 Execution
    let step1Res = '';
    try {
      step1Res = await runStepWithFallback(1, inputData, activeTab, apiKey, selectedModel, (model, status, err) => {
        setStep1(prev => ({ ...prev, model, status, error: err }));
      });
      const meta = parseJSONMetadata(step1Res);
      setStep1(prev => ({ ...prev, status: 'success', result: step1Res, metadata: meta }));
    } catch (err: any) {
      setStep1(prev => ({ ...prev, status: 'error', error: err.message }));
      setStep2(prev => ({ ...prev, status: 'error', error: 'Đã dừng do lỗi ở bước trước' }));
      setStep3(prev => ({ ...prev, status: 'error', error: 'Đã dừng do lỗi ở bước trước' }));
      setGlobalError(err.message);
      setLoading(false);
      return;
    }

    // Step 2 Execution
    setStep2({ status: 'loading', model: selectedModel, result: '' });
    let step2Res = '';
    try {
      step2Res = await runStepWithFallback(2, inputData, activeTab, apiKey, selectedModel, (model, status, err) => {
        setStep2(prev => ({ ...prev, model, status, error: err }));
      });
      const meta = parseJSONMetadata(step2Res);
      setStep2(prev => ({ ...prev, status: 'success', result: step2Res, metadata: meta }));
    } catch (err: any) {
      setStep2(prev => ({ ...prev, status: 'error', error: err.message }));
      setStep3(prev => ({ ...prev, status: 'error', error: 'Đã dừng do lỗi ở bước trước' }));
      setGlobalError(err.message);
      setLoading(false);
      return;
    }

    // Step 3 Execution
    setStep3({ status: 'loading', model: selectedModel, result: '' });
    let step3Res = '';
    try {
      step3Res = await runStepWithFallback(3, inputData, activeTab, apiKey, selectedModel, (model, status, err) => {
        setStep3(prev => ({ ...prev, model, status, error: err }));
      });
      const meta = parseJSONMetadata(step3Res);
      setStep3(prev => ({ ...prev, status: 'success', result: step3Res, metadata: meta }));
    } catch (err: any) {
      setStep3(prev => ({ ...prev, status: 'error', error: err.message }));
      setGlobalError(err.message);
      setLoading(false);
      return;
    }

    setLoading(false);
  };

  const runStepWithFallback = async (
    stepNumber: number,
    prompt: string,
    mode: string,
    userApiKey: string,
    initialModel: string,
    updateState: (model: string, status: 'loading' | 'error', error?: string) => void
  ): Promise<string> => {
    let startIndex = modelsChain.indexOf(initialModel);
    if (startIndex === -1) startIndex = 0;
    
    const modelsToTry = [
      ...modelsChain.slice(startIndex),
      ...modelsChain.slice(0, startIndex)
    ];
    
    let lastError = '';
    
    for (const modelId of modelsToTry) {
      // Build list of actual API models to try for this placeholder
      const actualModels: string[] = [];
      if (modelId === 'gemini-3-flash-preview') {
        actualModels.push('gemini-3-flash-preview', 'gemini-2.5-flash', 'gemini-2.0-flash');
      } else if (modelId === 'gemini-3-pro-preview') {
        actualModels.push('gemini-3-pro-preview', 'gemini-2.5-pro');
      } else if (modelId === 'gemini-2.5-flash') {
        actualModels.push('gemini-2.5-flash', 'gemini-2.0-flash');
      } else {
        actualModels.push(modelId);
      }

      for (const actualModel of actualModels) {
        // Try v1 first (stable), then v1beta
        for (const apiVersion of ['v1', 'v1beta']) {
          updateState(`${modelId} (${actualModel} - ${apiVersion})`, 'loading');
          
          const contextPrompt = `
Vai trò người dùng hiện tại đang là: ${mode} (Teacher, Student, Parent).
Đang thực hiện bước xử lý: BƯỚC ${stepNumber} (Step ${stepNumber}).

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

          try {
            const url = `https://generativelanguage.googleapis.com/${apiVersion}/models/${actualModel}:generateContent?key=${userApiKey}`;
            const response = await fetch(url, {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify({
                contents: [
                  {
                    parts: [{ text: contextPrompt }]
                  }
                ],
                systemInstruction: {
                  parts: [{ text: SYSTEM_INSTRUCTION }]
                },
                generationConfig: {
                  temperature: 0.7
                }
              }),
            });
            
            const data = await response.json();
            if (response.ok) {
              const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
              if (text) {
                // Successfully generated content!
                return text;
              }
              throw new Error("Không nhận được phản hồi hợp lệ từ Gemini API.");
            } else {
              lastError = data.error?.message || 'Lỗi không xác định từ Gemini API';
              console.warn(`Model ${actualModel} (${apiVersion}) failed for Step ${stepNumber}:`, lastError);
            }
          } catch (err: any) {
            lastError = err.message || 'Lỗi kết nối mạng';
            console.warn(`Model ${actualModel} (${apiVersion}) threw error for Step ${stepNumber}:`, lastError);
          }
        }
      }
    }
    
    let availableModelsStr = '';
    try {
      const listResponse = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${userApiKey}`);
      if (listResponse.ok) {
        const listData = await listResponse.json();
        const names = listData.models?.map((m: any) => m.name.replace('models/', '')) || [];
        if (names.length > 0) {
          availableModelsStr = '\n\n👉 Các model khả dụng cho API Key của bạn: ' + names.join(', ');
        }
      }
    } catch (e: any) {
      console.warn("Failed to list models:", e);
    }

    throw new Error((lastError || 'Tất cả các model và phiên bản API đều thất bại.') + availableModelsStr);
  };

  // Slide PowerPoint Export Helper
  const handleExportPPTX = (slideData: any[]) => {
    if (!slideData || !Array.isArray(slideData)) return;
    try {
      const pres = new pptxgen();
      pres.layout = 'LAYOUT_16x9';
      
      slideData.forEach((slideItem: any) => {
        const slide = pres.addSlide();
        
        // Solid dark blue accent header line
        slide.addShape(pres.ShapeType.rect, {
          x: 0.0, y: 0.0, w: 10.0, h: 0.15, fill: { color: '3B82F6' }
        });

        // Slide Title
        slide.addText(slideItem.title || 'Slide Title', {
          x: 0.6, y: 0.6, w: 8.8, h: 0.8,
          fontSize: 22, bold: true, color: '1E293B', fontFace: 'Calibri'
        });
        
        // Bullet points
        if (slideItem.bulletPoints && Array.isArray(slideItem.bulletPoints)) {
          const bullets = slideItem.bulletPoints.map((pt: string) => ({ text: pt, options: { bullet: true, indent: 20 } }));
          slide.addText(bullets, {
            x: 0.6, y: 1.6, w: 8.8, h: 3.8,
            fontSize: 14, color: '475569', fontFace: 'Calibri', paraSpaceBefore: 12
          });
        }
        
        // Speaker notes
        if (slideItem.notes) {
          slide.addNotes(slideItem.notes);
        }
      });
      
      pres.writeFile({ fileName: `EduVideo_Outline_Lecture.pptx` });
    } catch (err) {
      alert('Lỗi xuất file PowerPoint: ' + err);
    }
  };

  // Word Document Export Helper
  const handleExportDOCX = (worksheetData: any) => {
    if (!worksheetData) return;
    try {
      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              text: worksheetData.title || "Tài liệu học tập bổ trợ",
              heading: HeadingLevel.TITLE,
              spacing: { after: 300 }
            }),
            ...(worksheetData.exercises || []).flatMap((ex: any, idx: number) => [
              new Paragraph({
                children: [
                  new TextRun({ text: `Câu ${idx + 1}: `, bold: true }),
                  new TextRun({ text: ex.question || "" })
                ],
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 100 }
              }),
              ...(ex.options ? ex.options.map((opt: string) => 
                new Paragraph({ text: opt, indent: { left: 360 }, spacing: { after: 50 } })
              ) : []),
              new Paragraph({
                children: [
                  new TextRun({ text: "Gợi ý / Đáp án: ", bold: true, italic: true }),
                  new TextRun({ text: ex.answer || "Chưa cập nhật." })
                ],
                spacing: { before: 100, after: 150 }
              })
            ])
          ]
        }]
      });

      Packer.toBlob(doc).then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${worksheetData.title || 'EduVideo_TaiLieu'}.docx`;
        a.click();
      });
    } catch (err) {
      alert('Lỗi xuất file Word: ' + err);
    }
  };

  const tabs: { id: Role; label: string; icon: any; desc: string; placeholder: string }[] = [
    { 
      id: 'Teacher', 
      label: 'Teacher Portal', 
      icon: BookOpen,
      desc: 'Thiết kế bài giảng video, tích hợp câu hỏi trắc nghiệm & ma trận bài tập.',
      placeholder: 'Nhập chủ đề bài học, tài liệu tóm tắt hoặc nội dung chính để tạo kịch bản video giảng dạy chi tiết và câu hỏi tương tác...'
    },
    { 
      id: 'Student', 
      label: 'Student Portal', 
      icon: GraduationCap,
      desc: 'Tóm tắt bài học video, ghi chú chương mục & câu hỏi tự luyện tập phản xạ.',
      placeholder: 'Nhập nội dung/transcript của video hoặc câu hỏi bạn thắc mắc để tạo ghi chú thông minh và câu hỏi ôn tập phản xạ...'
    },
    { 
      id: 'Parent', 
      label: 'Parent Dashboard', 
      icon: Users,
      desc: 'Phân tích tiến độ học tập qua video, chẩn đoán lỗ hổng kiến thức & lộ trình ôn tập.',
      placeholder: 'Nhập tiến trình học tập của học sinh hoặc các bài quiz đã làm để nhận chẩn đoán lỗ hổng và lộ trình cải thiện 7 ngày...'
    }
  ];

  const currentTabInfo = tabs.find(t => t.id === activeTab)!;

  const getStepNames = (role: Role) => {
    switch(role) {
      case 'Teacher':
        return {
          step1: 'Video Script (Kịch bản)',
          step2: 'Video Quizzes (Câu hỏi)',
          step3: 'Supplemental Materials (Tài liệu)'
        };
      case 'Student':
        return {
          step1: 'Concept Map & Summary (Tóm tắt)',
          step2: 'Chapters & Notes (Ghi chú)',
          step3: 'Active Recall (Câu hỏi phản xạ)'
        };
      case 'Parent':
        return {
          step1: 'Learning Pulse (Tiến độ)',
          step2: 'Gap Analysis (Lỗ hổng kiến thức)',
          step3: 'Remedial Path (Lộ trình cải thiện)'
        };
    }
  };

  const stepNames = getStepNames(activeTab);

  const showBlockingModal = !apiKey.trim() && !isSettingsOpen;

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans">
      {/* Header */}
      <header className="bg-slate-950 border-b border-slate-800 px-6 py-4 flex items-center justify-between shadow-lg static top-0 z-40">
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-tr from-blue-600 to-indigo-600 p-2.5 rounded-xl shadow-md shadow-blue-900/40">
            <Video className="w-6 h-6 text-white" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-xl font-extrabold tracking-tight bg-gradient-to-r from-blue-400 via-indigo-400 to-purple-400 bg-clip-text text-transparent">EduVideo_AI</h1>
              <span className="text-[10px] uppercase font-bold tracking-widest px-2 py-0.5 bg-blue-500/20 text-blue-400 rounded-full border border-blue-500/30">V2.1</span>
            </div>
            <p className="text-xs text-slate-400 font-medium">AI-Powered Video Pedagogical Orchestrator</p>
          </div>
        </div>

        {/* API Key settings & warning */}
        <div className="flex items-center gap-3">
          {!apiKey && (
            <span className="hidden sm:inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-red-950/50 border border-red-500/30 text-[11px] font-semibold text-red-400 animate-pulse">
              <AlertCircle className="w-3.5 h-3.5 text-red-500" />
              Lấy API key để sử dụng app
            </span>
          )}
          <button
            onClick={() => {
              setTempKey(apiKey);
              setTempModel(selectedModel);
              setIsSettingsOpen(true);
            }}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl transition-all font-medium text-xs border
              ${!apiKey 
                ? 'bg-red-950/40 text-red-400 hover:bg-red-950/60 border-red-500/40' 
                : 'bg-slate-800 text-slate-200 hover:bg-slate-700 border-slate-700'
              }`}
          >
            <Settings className="w-4 h-4" />
            Cài đặt
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 w-full max-w-7xl mx-auto p-4 md:p-6 lg:p-8 flex flex-col lg:flex-row gap-6">
        
        {/* Sidebar */}
        <aside className="w-full lg:w-80 flex flex-col gap-6">
          {/* Navigation / Role Selector */}
          <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 p-3 flex flex-col gap-1.5 shadow-xl">
            <h3 className="text-[11px] font-bold text-slate-500 uppercase tracking-widest px-3 mb-1">Cổng điều phối</h3>
            {tabs.map((tab) => (
              <button
                key={tab.id}
                onClick={() => {
                  if (loading) return;
                  setActiveTab(tab.id);
                  setStep1({ status: 'idle', model: '', result: '' });
                  setStep2({ status: 'idle', model: '', result: '' });
                  setStep3({ status: 'idle', model: '', result: '' });
                  setGlobalError(null);
                  setPlayGame(false);
                }}
                disabled={loading}
                className={`flex items-center gap-3.5 px-4 py-3.5 rounded-xl transition-all font-semibold text-left text-sm disabled:opacity-50 disabled:cursor-not-allowed
                  ${activeTab === tab.id 
                    ? 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-md shadow-blue-500/20' 
                    : 'text-slate-400 hover:bg-slate-800/60 hover:text-slate-200'
                  }`}
              >
                <tab.icon className={`w-5 h-5 ${activeTab === tab.id ? 'text-white' : 'text-slate-500'}`} />
                <div>
                  <p className="font-bold leading-tight">{tab.label}</p>
                  <p className={`text-[10px] font-normal mt-0.5 max-w-[200px] truncate ${activeTab === tab.id ? 'text-blue-100' : 'text-slate-500'}`}>{tab.desc}</p>
                </div>
              </button>
            ))}
          </div>

          {/* Model Stats info */}
          <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 p-5 flex flex-col gap-4 shadow-xl">
            <div className="flex items-center gap-2 pb-3 border-b border-slate-800">
              <Sparkles className="w-4 h-4 text-blue-400" />
              <h4 className="text-xs font-bold uppercase tracking-wider text-slate-300">Thông tin cấu hình</h4>
            </div>
            
            <div className="flex flex-col gap-3 text-xs">
              <div>
                <span className="text-slate-500 block mb-1">Model ưu tiên:</span>
                <span className="font-mono text-blue-400 font-bold bg-blue-950/40 border border-blue-900/50 px-2.5 py-1 rounded-md inline-block">
                  {selectedModel}
                </span>
              </div>

              <div>
                <span className="text-slate-500 block mb-1">Hàng đợi Fallback (Lỗi tự chuyển):</span>
                <div className="flex flex-wrap items-center gap-1.5 font-mono text-[10px] text-slate-400">
                  <span className="px-1.5 py-0.5 bg-slate-800/50 rounded">flash-3</span>
                  <span className="text-slate-600">&rarr;</span>
                  <span className="px-1.5 py-0.5 bg-slate-800/50 rounded">pro-3</span>
                  <span className="text-slate-600">&rarr;</span>
                  <span className="px-1.5 py-0.5 bg-slate-800/50 rounded">flash-2.5</span>
                  <span className="text-slate-600">&rarr;</span>
                  <span className="px-1.5 py-0.5 bg-slate-800/50 rounded">pro-2.5</span>
                </div>
              </div>

              <div>
                <span className="text-slate-500 block mb-1">Trạng thái API Key:</span>
                {apiKey ? (
                  <span className="text-emerald-400 font-semibold flex items-center gap-1">
                    <CheckCircle2 className="w-3.5 h-3.5" /> Sẵn sàng (Hoạt động)
                  </span>
                ) : (
                  <span className="text-red-400 font-semibold flex items-center gap-1">
                    <XCircle className="w-3.5 h-3.5 animate-pulse" /> Chưa thiết lập API key
                  </span>
                )}
              </div>
            </div>
          </div>
        </aside>

        {/* Workspace Area */}
        <section className="flex-1 flex flex-col gap-6">
          
          {/* Input Area */}
          <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 p-6 flex flex-col gap-4 shadow-xl">
            <div className="flex items-center justify-between">
              <h2 className="text-base font-bold text-slate-200 flex items-center gap-2">
                <currentTabInfo.icon className="w-5 h-5 text-blue-500" />
                {currentTabInfo.label} Input
              </h2>
              <span className="text-xs text-slate-500 font-medium">{activeTab} Mode</span>
            </div>
            
            <textarea
              className="w-full h-36 p-4 bg-slate-900/60 border border-slate-800 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 resize-none text-slate-300 placeholder-slate-500 text-sm leading-relaxed transition-all"
              placeholder={currentTabInfo.placeholder}
              value={inputData}
              onChange={(e) => setInputData(e.target.value)}
              disabled={loading}
            />

            <div className="flex justify-between items-center mt-1">
              <div className="text-xs text-slate-500 font-medium">
                Sử dụng tiếng Việt hoặc ngôn ngữ của tài liệu nguồn.
              </div>
              <button
                onClick={handleGenerate}
                disabled={loading || !inputData.trim()}
                className="flex items-center gap-2 px-6 py-3 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl disabled:opacity-40 disabled:cursor-not-allowed font-semibold text-xs shadow-md shadow-blue-500/10 hover:shadow-blue-500/20 transition-all"
              >
                {loading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Đang xử lý đa nhiệm...
                  </>
                ) : (
                  <>
                    <Send className="w-4 h-4" />
                    Bắt đầu quy trình EduVideo
                  </>
                )}
              </button>
            </div>
          </div>

          {/* Stepper Progress Board */}
          {(step1.status !== 'idle' || step2.status !== 'idle' || step3.status !== 'idle') && (
            <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 p-5 shadow-xl">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-4 flex items-center justify-between">
                <span>Tiến trình xử lý 3 bước (AI Multi-step Pipeline)</span>
                {loading && <span className="text-[10px] px-2 py-0.5 bg-blue-500/15 text-blue-400 rounded-md border border-blue-500/20 animate-pulse">Running</span>}
              </h3>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {/* Step 1 Card */}
                <div className={`p-3.5 rounded-xl border transition-all ${
                  step1.status === 'loading' ? 'bg-blue-950/20 border-blue-500/40 ring-1 ring-blue-500/20' :
                  step1.status === 'success' ? 'bg-emerald-950/10 border-emerald-500/30' :
                  step1.status === 'error' ? 'bg-red-950/10 border-red-500/30' :
                  'bg-slate-900/40 border-slate-800/80'
                }`}>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-slate-500 uppercase">Bước 1</span>
                    {step1.status === 'loading' && <Loader2 className="w-4 h-4 text-blue-500 animate-spin" />}
                    {step1.status === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                    {step1.status === 'error' && <XCircle className="w-4 h-4 text-red-500" />}
                    {step1.status === 'waiting' && <span className="w-2 h-2 rounded-full bg-slate-700"></span>}
                  </div>
                  <h4 className="text-xs font-bold text-slate-200 truncate">{stepNames.step1}</h4>
                  <div className="mt-2 flex flex-col gap-1">
                    <span className="text-[10px] text-slate-400">
                      Trạng thái: 
                      <span className={`font-semibold ml-1 ${
                        step1.status === 'loading' ? 'text-blue-400' :
                        step1.status === 'success' ? 'text-emerald-400' :
                        step1.status === 'error' ? 'text-red-400' : 'text-slate-500'
                      }`}>
                        {step1.status === 'waiting' && 'Chờ xử lý'}
                        {step1.status === 'loading' && 'Đang xử lý...'}
                        {step1.status === 'success' && 'Hoàn tất'}
                        {step1.status === 'error' && 'Đã dừng do lỗi'}
                      </span>
                    </span>
                    {step1.model && (
                      <span className="text-[9px] font-mono text-slate-500">
                        Model: {step1.model}
                      </span>
                    )}
                  </div>
                </div>

                {/* Step 2 Card */}
                <div className={`p-3.5 rounded-xl border transition-all ${
                  step2.status === 'loading' ? 'bg-blue-950/20 border-blue-500/40 ring-1 ring-blue-500/20' :
                  step2.status === 'success' ? 'bg-emerald-950/10 border-emerald-500/30' :
                  step2.status === 'error' ? 'bg-red-950/10 border-red-500/30' :
                  'bg-slate-900/40 border-slate-800/80'
                }`}>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-slate-500 uppercase">Bước 2</span>
                    {step2.status === 'loading' && <Loader2 className="w-4 h-4 text-blue-500 animate-spin" />}
                    {step2.status === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                    {step2.status === 'error' && <XCircle className="w-4 h-4 text-red-500" />}
                    {step2.status === 'waiting' && <span className="w-2 h-2 rounded-full bg-slate-700"></span>}
                  </div>
                  <h4 className="text-xs font-bold text-slate-200 truncate">{stepNames.step2}</h4>
                  <div className="mt-2 flex flex-col gap-1">
                    <span className="text-[10px] text-slate-400">
                      Trạng thái: 
                      <span className={`font-semibold ml-1 ${
                        step2.status === 'loading' ? 'text-blue-400' :
                        step2.status === 'success' ? 'text-emerald-400' :
                        step2.status === 'error' ? 'text-red-400' : 'text-slate-500'
                      }`}>
                        {step2.status === 'waiting' && 'Chờ xử lý'}
                        {step2.status === 'loading' && 'Đang xử lý...'}
                        {step2.status === 'success' && 'Hoàn tất'}
                        {step2.status === 'error' && 'Đã dừng do lỗi'}
                      </span>
                    </span>
                    {step2.model && (
                      <span className="text-[9px] font-mono text-slate-500">
                        Model: {step2.model}
                      </span>
                    )}
                  </div>
                </div>

                {/* Step 3 Card */}
                <div className={`p-3.5 rounded-xl border transition-all ${
                  step3.status === 'loading' ? 'bg-blue-950/20 border-blue-500/40 ring-1 ring-blue-500/20' :
                  step3.status === 'success' ? 'bg-emerald-950/10 border-emerald-500/30' :
                  step3.status === 'error' ? 'bg-red-950/10 border-red-500/30' :
                  'bg-slate-900/40 border-slate-800/80'
                }`}>
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-slate-500 uppercase">Bước 3</span>
                    {step3.status === 'loading' && <Loader2 className="w-4 h-4 text-blue-500 animate-spin" />}
                    {step3.status === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                    {step3.status === 'error' && <XCircle className="w-4 h-4 text-red-500" />}
                    {step3.status === 'waiting' && <span className="w-2 h-2 rounded-full bg-slate-700"></span>}
                  </div>
                  <h4 className="text-xs font-bold text-slate-200 truncate">{stepNames.step3}</h4>
                  <div className="mt-2 flex flex-col gap-1">
                    <span className="text-[10px] text-slate-400">
                      Trạng thái: 
                      <span className={`font-semibold ml-1 ${
                        step3.status === 'loading' ? 'text-blue-400' :
                        step3.status === 'success' ? 'text-emerald-400' :
                        step3.status === 'error' ? 'text-red-400' : 'text-slate-500'
                      }`}>
                        {step3.status === 'waiting' && 'Chờ xử lý'}
                        {step3.status === 'loading' && 'Đang xử lý...'}
                        {step3.status === 'success' && 'Hoàn tất'}
                        {step3.status === 'error' && 'Đã dừng do lỗi'}
                      </span>
                    </span>
                    {step3.model && (
                      <span className="text-[9px] font-mono text-slate-500">
                        Model: {step3.model}
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {globalError && (
                <div className="mt-4 p-4 bg-red-950/30 border border-red-900/50 rounded-xl flex gap-3 text-red-300 items-start">
                  <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                  <div>
                    <h4 className="text-xs font-bold text-red-400">Thông báo lỗi từ API (Nguyên văn):</h4>
                    <p className="text-xs font-mono mt-1 select-all break-all">{globalError}</p>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Output Display Areas */}
          {(step1.result || step2.result || step3.result) && (
            <div className="flex flex-col gap-6">
              
              {/* Step 1 Output */}
              {step1.result && (
                <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 overflow-hidden shadow-xl">
                  <div className="bg-gradient-to-r from-blue-900/30 to-indigo-900/20 px-6 py-4 border-b border-slate-800 flex justify-between items-center">
                    <span className="text-[11px] font-bold text-blue-400 uppercase tracking-widest">Bước 1: {stepNames.step1}</span>
                    <div className="flex items-center gap-2">
                      {activeTab === 'Teacher' && step1.metadata?.slideData && (
                        <button 
                          onClick={() => handleExportPPTX(step1.metadata.slideData)}
                          className="flex items-center gap-1.5 px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded text-[10px] font-bold transition-all"
                        >
                          <Download className="w-3 h-3" />
                          Xuất PPTX (.pptx)
                        </button>
                      )}
                      <span className="text-[9px] font-mono px-2 py-0.5 bg-slate-800 text-slate-400 rounded border border-slate-700">{step1.model}</span>
                    </div>
                  </div>
                  <div className="p-6 md:p-8">
                    <div className="prose prose-invert prose-blue max-w-none prose-headings:font-bold prose-headings:text-slate-100 prose-headings:tracking-tight prose-p:text-slate-300 prose-p:leading-relaxed prose-a:text-blue-400 prose-strong:text-slate-100">
                      <Markdown>{cleanMarkdownText(step1.result)}</Markdown>
                    </div>
                    {/* SVG Concept Map Visualization for Student Step 1 */}
                    {activeTab === 'Student' && step1.metadata?.conceptMapData && (
                      <ConceptMapGraph data={step1.metadata.conceptMapData} />
                    )}
                    {/* SVG Learning Pulse for Parent Step 1 */}
                    {activeTab === 'Parent' && step1.metadata?.pulseData && (
                      <LearningPulseChart data={step1.metadata.pulseData} />
                    )}
                  </div>
                </div>
              )}

              {/* Step 2 Output */}
              {step2.result && (
                <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 overflow-hidden shadow-xl">
                  <div className="bg-gradient-to-r from-indigo-900/30 to-purple-900/20 px-6 py-4 border-b border-slate-800 flex justify-between items-center">
                    <span className="text-[11px] font-bold text-indigo-400 uppercase tracking-widest">Bước 2: {stepNames.step2}</span>
                    <span className="text-[9px] font-mono px-2 py-0.5 bg-slate-800 text-slate-400 rounded border border-slate-700">{step2.model}</span>
                  </div>
                  <div className="p-6 md:p-8">
                    <div className="prose prose-invert prose-blue max-w-none prose-headings:font-bold prose-headings:text-slate-100 prose-headings:tracking-tight prose-p:text-slate-300 prose-p:leading-relaxed prose-a:text-indigo-400 prose-strong:text-slate-100">
                      <Markdown>{cleanMarkdownText(step2.result)}</Markdown>
                    </div>
                    {/* SVG Radar Chart for Parent Step 2 */}
                    {activeTab === 'Parent' && step2.metadata?.gapData && (
                      <GapAnalysisRadarChart data={step2.metadata.gapData} />
                    )}
                  </div>
                </div>
              )}

              {/* Step 3 Output */}
              {step3.result && (
                <div className="bg-slate-950/65 backdrop-blur-md rounded-2xl border border-slate-800/80 overflow-hidden shadow-xl">
                  <div className="bg-gradient-to-r from-purple-900/30 to-pink-900/20 px-6 py-4 border-b border-slate-800 flex justify-between items-center">
                    <span className="text-[11px] font-bold text-purple-400 uppercase tracking-widest">Bước 3: {stepNames.step3}</span>
                    <div className="flex items-center gap-2">
                      {activeTab === 'Teacher' && step3.metadata?.worksheetData && (
                        <button 
                          onClick={() => handleExportDOCX(step3.metadata.worksheetData)}
                          className="flex items-center gap-1.5 px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white rounded text-[10px] font-bold transition-all"
                        >
                          <Download className="w-3 h-3" />
                          Xuất Đề thi (.docx)
                        </button>
                      )}
                      {activeTab === 'Student' && step3.metadata?.quizData && (
                        <button 
                          onClick={() => setPlayGame(!playGame)}
                          className="flex items-center gap-1.5 px-2.5 py-1 bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-500 hover:to-orange-500 text-white rounded text-[10px] font-bold transition-all"
                        >
                          <Flame className="w-3 h-3" />
                          {playGame ? 'Ẩn Game ôn tập' : 'Chơi Game phản xạ'}
                        </button>
                      )}
                      <span className="text-[9px] font-mono px-2 py-0.5 bg-slate-800 text-slate-400 rounded border border-slate-700">{step3.model}</span>
                    </div>
                  </div>
                  <div className="p-6 md:p-8">
                    {/* Game Mode toggle for Student Step 3 */}
                    {activeTab === 'Student' && playGame && step3.metadata?.quizData ? (
                      <ActiveRecallGame data={step3.metadata.quizData} />
                    ) : (
                      <div className="prose prose-invert prose-blue max-w-none prose-headings:font-bold prose-headings:text-slate-100 prose-headings:tracking-tight prose-p:text-slate-300 prose-p:leading-relaxed prose-a:text-purple-400 prose-strong:text-slate-100">
                        <Markdown>{cleanMarkdownText(step3.result)}</Markdown>
                      </div>
                    )}
                  </div>
                </div>
              )}

            </div>
          )}

          {/* Empty State */}
          {!loading && !step1.result && !step2.result && !step3.result && (
            <div className="flex-1 flex flex-col items-center justify-center py-12 px-4 border border-dashed border-slate-800 rounded-2xl bg-slate-950/20 text-center">
              <div className="w-12 h-12 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center text-slate-500 mb-4 shadow-inner">
                <Play className="w-5 h-5 ml-0.5" />
              </div>
              <h3 className="text-sm font-bold text-slate-300">Không có kịch bản hoạt động</h3>
              <p className="text-xs text-slate-500 mt-1 max-w-md">
                Thiết lập API Key (nếu chưa có) và chọn cổng tương tác bên trái để bắt đầu thiết kế kịch bản video giảng dạy.
              </p>
            </div>
          )}

        </section>
      </main>

      {/* Settings Modal */}
      {(isSettingsOpen || showBlockingModal) && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/80 backdrop-blur-md p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full overflow-hidden shadow-2xl animate-fade-in flex flex-col">
            <div className="bg-slate-950 px-6 py-5 border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-2">
                <Key className="w-4 h-4 text-blue-500" />
                <h2 className="text-sm font-bold text-slate-200">
                  {showBlockingModal ? 'Yêu cầu Cấu hình API Key Gemini' : 'Thiết lập EduVideo_AI'}
                </h2>
              </div>
              {!showBlockingModal && (
                <button 
                  onClick={() => setIsSettingsOpen(false)}
                  className="text-slate-500 hover:text-slate-300 text-xs font-semibold p-1 hover:bg-slate-800 rounded-lg transition-colors"
                >
                  Đóng
                </button>
              )}
            </div>

            <div className="p-6 flex flex-col gap-6">
              {showBlockingModal && (
                <div className="p-4 bg-red-950/40 border border-red-500/20 text-red-300 rounded-xl text-xs flex gap-3">
                  <AlertCircle className="w-5 h-5 text-red-400 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-bold">Ứng dụng chưa có API Key!</span> Vui lòng nhập API Key Gemini cá nhân bên dưới để tiếp tục trải nghiệm đầy đủ các tính năng của EduVideo_AI.
                  </div>
                </div>
              )}

              {/* API Key Input */}
              <div className="flex flex-col gap-2">
                <label className="text-xs font-bold text-slate-400 flex items-center justify-between">
                  <span>API Key Gemini:</span>
                  <a 
                    href="https://aistudio.google.com/api-keys" 
                    target="_blank" 
                    rel="noreferrer"
                    className="text-[10px] text-red-400 hover:underline flex items-center gap-1 font-semibold"
                  >
                    Lấy API key tại Google AI Studio <ExternalLink className="w-3 h-3" />
                  </a>
                </label>
                <input
                  type="password"
                  placeholder="Nhập API Key của bạn (AIzaSy...)"
                  className="w-full px-4 py-3 bg-slate-950 border border-slate-800 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 text-xs font-mono text-slate-300"
                  value={tempKey}
                  onChange={(e) => setTempKey(e.target.value)}
                />
              </div>

              {/* Model selection cards */}
              <div className="flex flex-col gap-3">
                <span className="text-xs font-bold text-slate-400">Chọn Model AI ưu tiên (Có cơ chế tự động Fallback):</span>
                <div className="grid grid-cols-1 gap-2.5">
                  {[
                    { id: 'gemini-3-flash-preview', title: 'gemini-3-flash-preview', badge: 'Mặc định', desc: 'Tốc độ cực nhanh, tối ưu hóa chi phí và các tác vụ kịch bản.' },
                    { id: 'gemini-3-pro-preview', title: 'gemini-3-pro-preview', badge: 'Hiệu năng cao', desc: 'Thông minh vượt trội, thích hợp cho các bài toán phân tích lỗ hổng phức tạp.' },
                    { id: 'gemini-2.5-flash', title: 'gemini-2.5-flash', badge: 'Mạnh mẽ', desc: 'Độ chính xác cao, xử lý văn cảnh sư phạm linh hoạt.' }
                  ].map((modelOpt) => (
                    <button
                      key={modelOpt.id}
                      onClick={() => setTempModel(modelOpt.id)}
                      className={`flex flex-col text-left p-3 rounded-xl border transition-all cursor-pointer
                        ${tempModel === modelOpt.id 
                          ? 'bg-blue-950/30 border-blue-500/70 ring-1 ring-blue-500/30' 
                          : 'bg-slate-950/40 border-slate-800 hover:bg-slate-800/40'
                        }`}
                    >
                      <div className="flex items-center justify-between mb-1.5">
                        <span className="text-xs font-bold font-mono text-slate-200">{modelOpt.title}</span>
                        <span className={`text-[9px] font-semibold px-2 py-0.5 rounded-full ${
                          modelOpt.badge === 'Mặc định' ? 'bg-blue-500/20 text-blue-400' :
                          modelOpt.badge === 'Hiệu năng cao' ? 'bg-purple-500/20 text-purple-400' :
                          'bg-slate-800 text-slate-400'
                        }`}>{modelOpt.badge}</span>
                      </div>
                      <p className="text-[10px] text-slate-400 leading-normal">{modelOpt.desc}</p>
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <div className="bg-slate-950 px-6 py-4 border-t border-slate-800 flex justify-end gap-2">
              <button
                onClick={() => saveSettings(tempKey, tempModel)}
                disabled={showBlockingModal && !tempKey.trim()}
                className="px-6 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl text-xs font-bold disabled:opacity-40 disabled:cursor-not-allowed shadow-md shadow-blue-500/10 transition-all"
              >
                Lưu cài đặt
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ==================== VISUALIZATION COMPONENTS ====================

// SVG Interactive Concept Map Component
function ConceptMapGraph({ data }: { data: any }) {
  if (!data || !data.nodes) return null;
  const width = 500;
  const height = 240;
  
  const nodes = data.nodes.map((node: any, idx: number) => {
    let x = width / 2;
    let y = height / 2;
    if (node.type === 'sub') {
      const subNodes = data.nodes.filter((n: any) => n.type === 'sub');
      const subIdx = subNodes.findIndex((n: any) => n.id === node.id);
      const angle = (subIdx * 2 * Math.PI) / (subNodes.length || 1);
      x = width / 2 + 160 * Math.cos(angle);
      y = height / 2 + 65 * Math.sin(angle);
    }
    return { ...node, x, y };
  });

  const findNodeCoords = (id: string) => {
    const n = nodes.find((node: any) => node.id === id);
    return n ? { x: n.x, y: n.y } : { x: width/2, y: height/2 };
  };

  const [hoveredNode, setHoveredNode] = useState<string | null>(null);

  return (
    <div className="bg-slate-900 border border-slate-800 p-4 rounded-xl shadow-inner mt-6 flex flex-col items-center">
      <div className="flex justify-between items-center w-full mb-3">
        <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
          <Sparkles className="w-3.5 h-3.5 text-indigo-400 animate-pulse" />
          Bản đồ khái niệm tương tác (Concept Map)
        </h4>
        <span className="text-[10px] bg-indigo-500/10 text-indigo-400 px-2 py-0.5 rounded border border-indigo-500/20">Hover Node</span>
      </div>

      <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-auto overflow-visible">
        {/* Draw Links */}
        {data.links && data.links.map((link: any, idx: number) => {
          const src = findNodeCoords(link.source);
          const tgt = findNodeCoords(link.target);
          const isHighlighted = hoveredNode === link.source || hoveredNode === link.target;
          return (
            <line 
              key={idx} 
              x1={src.x} 
              y1={src.y} 
              x2={tgt.x} 
              y2={tgt.y} 
              stroke={isHighlighted ? "#6366f1" : "#334155"} 
              strokeWidth={isHighlighted ? 2.5 : 1}
              strokeDasharray={isHighlighted ? "none" : "3,3"}
              className="transition-all duration-300"
            />
          );
        })}

        {/* Draw Nodes */}
        {nodes.map((node: any, idx: number) => {
          const isMain = node.type === 'main';
          const isHovered = hoveredNode === node.id;
          return (
            <g 
              key={idx} 
              onMouseEnter={() => setHoveredNode(node.id)}
              onMouseLeave={() => setHoveredNode(null)}
              className="cursor-pointer"
            >
              <circle 
                cx={node.x} 
                cy={node.y} 
                r={isMain ? 22 : 14} 
                fill={isMain ? "#312e81" : "#0f172a"} 
                stroke={isHovered ? "#818cf8" : (isMain ? "#4f46e5" : "#475569")} 
                strokeWidth="2" 
                className="transition-all duration-300"
              />
              <text 
                x={node.x} 
                y={node.y + (isMain ? 34 : 24)} 
                fill={isHovered ? "#ffffff" : "#94a3b8"} 
                fontSize={isMain ? "9" : "8"} 
                fontWeight={isMain ? "bold" : "normal"}
                textAnchor="middle"
                className="select-none"
              >
                {node.label}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}

// SVG Interactive Learning Pulse Chart (Line & Area)
function LearningPulseChart({ data }: { data: any[] }) {
  if (!data || !Array.isArray(data) || data.length === 0) return null;
  const width = 500;
  const height = 200;
  const padding = 35;
  
  const points = data.map((d, i) => {
    const x = padding + (i * (width - 2 * padding)) / (data.length - 1);
    const y = height - padding - (d.watchPercent / 100) * (height - 2 * padding);
    return { x, y, day: d.day, val: d.watchPercent, score: d.score };
  });

  const pathD = points.reduce((acc, p, i) => i === 0 ? `M ${p.x} ${p.y}` : `${acc} L ${p.x} ${p.y}`, '');
  const areaD = `${pathD} L ${points[points.length - 1].x} ${height - padding} L ${points[0].x} ${height - padding} Z`;
  
  const [hoveredPoint, setHoveredPoint] = useState<number | null>(null);

  return (
    <div className="bg-slate-900 border border-slate-800 p-4 rounded-xl shadow-inner mt-6">
      <div className="flex justify-between items-center mb-4">
        <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
          <Award className="w-3.5 h-3.5 text-blue-400" />
          Nhịp độ học tập qua video (Tỉ lệ hoàn thành)
        </h4>
        <span className="text-[10px] bg-blue-500/10 text-blue-400 px-2 py-0.5 rounded border border-blue-500/20">Hover điểm số</span>
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-auto overflow-visible">
        <defs>
          <linearGradient id="areaGrad" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#3b82f6" stopOpacity="0.4" />
            <stop offset="100%" stopColor="#3b82f6" stopOpacity="0.0" />
          </linearGradient>
        </defs>
        {/* Grid lines */}
        {[0, 25, 50, 75, 100].map((grid) => {
          const y = height - padding - (grid / 100) * (height - 2 * padding);
          return (
            <g key={grid}>
              <line x1={padding} y1={y} x2={width - padding} y2={y} stroke="#334155" strokeDasharray="3,3" strokeWidth="0.5" />
              <text x={padding - 8} y={y + 3} fill="#64748b" fontSize="8" textAnchor="end">{grid}%</text>
            </g>
          );
        })}
        {/* Area */}
        <path d={areaD} fill="url(#areaGrad)" />
        {/* Line */}
        <path d={pathD} fill="none" stroke="#3b82f6" strokeWidth="2" strokeLinecap="round" />
        {/* Interactive Points */}
        {points.map((p, i) => (
          <g 
            key={i} 
            onMouseEnter={() => setHoveredPoint(i)}
            onMouseLeave={() => setHoveredPoint(null)}
            className="cursor-pointer"
          >
            <circle cx={p.x} cy={p.y} r={hoveredPoint === i ? 6 : 4} fill="#0f172a" stroke="#3b82f6" strokeWidth="2" />
            <text x={p.x} y={height - 10} fill="#64748b" fontSize="8" textAnchor="middle" className="select-none">{p.day}</text>
          </g>
        ))}
        {/* Tooltip */}
        {hoveredPoint !== null && (
          <g>
            <rect 
              x={Math.min(width - 105, Math.max(15, points[hoveredPoint].x - 50))} 
              y={points[hoveredPoint].y - 38} 
              width="100" 
              height="28" 
              rx="4" 
              fill="#0f172a" 
              stroke="#475569" 
              strokeWidth="1" 
            />
            <text 
              x={Math.min(width - 105, Math.max(15, points[hoveredPoint].x - 50)) + 50} 
              y={points[hoveredPoint].y - 25} 
              fill="#f1f5f9" 
              fontSize="8" 
              fontWeight="bold" 
              textAnchor="middle"
            >
              Đã xem: {points[hoveredPoint].val}%
            </text>
            <text 
              x={Math.min(width - 105, Math.max(15, points[hoveredPoint].x - 50)) + 50} 
              y={points[hoveredPoint].y - 15} 
              fill="#fbbf24" 
              fontSize="8" 
              fontWeight="bold" 
              textAnchor="middle"
            >
              Điểm quiz: {points[hoveredPoint].score}/100
            </text>
          </g>
        )}
      </svg>
    </div>
  );
}

// SVG Interactive Radar Chart (Gap Analysis)
function GapAnalysisRadarChart({ data }: { data: any[] }) {
  if (!data || !Array.isArray(data) || data.length === 0) return null;
  const width = 400;
  const height = 280;
  const cx = width / 2;
  const cy = height / 2;
  const r = 90;
  const numAxes = data.length;
  
  const getCoords = (i: number, val: number) => {
    const angle = (i * 2 * Math.PI) / numAxes - Math.PI / 2;
    const x = cx + r * (val / 100) * Math.cos(angle);
    const y = cy + r * (val / 100) * Math.sin(angle);
    return { x, y };
  };

  const gridPolygons = [20, 40, 60, 80, 100].map((gridVal) => {
    return data.map((_, idx) => {
      const { x, y } = getCoords(idx, gridVal);
      return `${x},${y}`;
    }).join(' ');
  });

  const dataPointsStr = data.map((d, idx) => {
    const { x, y } = getCoords(idx, d.score);
    return `${x},${y}`;
  }).join(' ');

  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null);

  return (
    <div className="bg-slate-900 border border-slate-800 p-4 rounded-xl shadow-inner mt-6 flex flex-col items-center">
      <div className="flex justify-between items-center w-full mb-3">
        <h4 className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
          <Settings className="w-3.5 h-3.5 text-purple-400" />
          Biểu đồ Radar chẩn đoán kỹ năng học tập
        </h4>
        <span className="text-[10px] bg-purple-500/10 text-purple-400 px-2 py-0.5 rounded border border-purple-500/20">Hover kỹ năng</span>
      </div>

      <svg viewBox={`0 0 ${width} ${height}`} className="w-full max-w-sm h-auto overflow-visible">
        {/* Web grids */}
        {gridPolygons.map((pts, idx) => (
          <polygon 
            key={idx} 
            points={pts} 
            fill="none" 
            stroke="#334155" 
            strokeWidth="0.5" 
            strokeDasharray="2,2" 
          />
        ))}
        {/* Grid axes */}
        {data.map((_, idx) => {
          const outer = getCoords(idx, 100);
          return (
            <line 
              key={idx} 
              x1={cx} 
              y1={cy} 
              x2={outer.x} 
              y2={outer.y} 
              stroke="#334155" 
              strokeWidth="0.5" 
            />
          );
        })}
        {/* Data Polygon Fill */}
        <polygon 
          points={dataPointsStr} 
          fill="#a855f7" 
          fillOpacity="0.25" 
          stroke="#c084fc" 
          strokeWidth="2" 
        />
        {/* Labels & Markers */}
        {data.map((d, idx) => {
          const outer = getCoords(idx, 115);
          const point = getCoords(idx, d.score);
          
          let textAnchor = "middle";
          if (outer.x < cx - 10) textAnchor = "end";
          if (outer.x > cx + 10) textAnchor = "start";
          
          return (
            <g key={idx}>
              <text 
                x={outer.x} 
                y={outer.y + 3} 
                fill={hoveredIndex === idx ? "#e9d5ff" : "#94a3b8"} 
                fontSize="8" 
                fontWeight={hoveredIndex === idx ? "bold" : "normal"}
                textAnchor={textAnchor}
                onMouseEnter={() => setHoveredIndex(idx)}
                onMouseLeave={() => setHoveredIndex(null)}
                className="cursor-pointer select-none"
              >
                {d.subject} ({d.score}%)
              </text>
              <circle 
                cx={point.x} 
                cy={point.y} 
                r={hoveredIndex === idx ? 5.5 : 3.5} 
                fill="#0f172a" 
                stroke="#c084fc" 
                strokeWidth="2" 
                onMouseEnter={() => setHoveredIndex(idx)}
                onMouseLeave={() => setHoveredIndex(null)}
                className="cursor-pointer transition-all duration-150"
              />
            </g>
          );
        })}
      </svg>
    </div>
  );
}

// ==================== GAMIFICATION COMPONENT ====================

// Interactive Flashcard Active Recall game
function ActiveRecallGame({ data }: { data: any[] }) {
  const [currentIndex, setCurrentIndex] = useState(0);
  const [isFlipped, setIsFlipped] = useState(false);
  const [hintLevel, setHintLevel] = useState(0);
  const [score, setScore] = useState(0);
  const [answeredCount, setAnsweredCount] = useState<number[]>([]);

  if (!data || !Array.isArray(data) || data.length === 0) return null;
  const currentCard = data[currentIndex];

  const handleNext = () => {
    setIsFlipped(false);
    setHintLevel(0);
    setCurrentIndex((prev) => (prev + 1) % data.length);
  };

  const handlePrev = () => {
    setIsFlipped(false);
    setHintLevel(0);
    setCurrentIndex((prev) => (prev - 1 + data.length) % data.length);
  };

  const handleCorrect = () => {
    if (!answeredCount.includes(currentIndex)) {
      setScore(prev => prev + 10);
      setAnsweredCount(prev => [...prev, currentIndex]);
    }
    setIsFlipped(true);
  };

  return (
    <div className="bg-slate-950 border border-slate-800 rounded-2xl p-5 shadow-2xl mt-4 max-w-lg mx-auto flex flex-col gap-4 animate-fade-in">
      <div className="flex justify-between items-center">
        <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest flex items-center gap-1.5">
          <Flame className="w-3.5 h-3.5 text-orange-500" />
          Game ôn luyện phản xạ (Thẻ {currentIndex + 1}/{data.length})
        </span>
        <span className="text-xs font-bold text-amber-400 bg-amber-950/40 border border-amber-900/50 px-2.5 py-1 rounded-lg">
          Score: {score} pts
        </span>
      </div>

      {/* Flippable Card */}
      <div 
        onClick={() => setIsFlipped(!isFlipped)}
        className="w-full h-48 bg-slate-900 border border-slate-800 rounded-xl flex items-center justify-center p-6 text-center cursor-pointer select-none relative overflow-hidden transition-all duration-300 transform hover:border-indigo-500/50"
      >
        <div className={`absolute inset-0 bg-gradient-to-tr transition-opacity duration-300 ${
          isFlipped ? 'from-purple-950/20 to-indigo-950/20 opacity-100' : 'from-slate-900/10 to-slate-950/10 opacity-30'
        }`} />
        
        {!isFlipped ? (
          <div className="z-10 flex flex-col gap-2">
            <span className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Câu hỏi ôn tập</span>
            <p className="text-sm font-bold text-slate-200">{currentCard.question}</p>
            <span className="text-[9px] text-slate-600 mt-2 italic">Click vào thẻ để xem đáp án ẩn</span>
          </div>
        ) : (
          <div className="z-10 flex flex-col gap-2">
            <span className="text-[10px] text-purple-400 uppercase tracking-wider font-bold">Đáp án chi tiết</span>
            <p className="text-sm font-semibold text-slate-200">{currentCard.answer}</p>
            <span className="text-[9px] text-slate-600 mt-2 italic">Click vào thẻ để úp lại</span>
          </div>
        )}
      </div>

      {/* Hints System */}
      {!isFlipped && currentCard.hints && currentCard.hints.length > 0 && (
        <div className="flex flex-col gap-2 bg-slate-900/50 border border-slate-800/60 p-3 rounded-xl">
          <div className="flex justify-between items-center">
            <span className="text-[10px] font-bold text-slate-400 uppercase">Hệ thống gợi ý thông minh</span>
            <button 
              onClick={(e) => { e.stopPropagation(); setHintLevel(prev => Math.min(prev + 1, currentCard.hints.length)); }}
              disabled={hintLevel >= currentCard.hints.length}
              className="text-[9px] font-semibold text-blue-400 hover:underline disabled:text-slate-600 disabled:no-underline"
            >
              Yêu cầu gợi ý (Cấp độ {hintLevel}/{currentCard.hints.length})
            </button>
          </div>
          {hintLevel > 0 && (
            <div className="text-xs text-blue-300 font-medium leading-relaxed mt-1">
              💡 {currentCard.hints.slice(0, hintLevel).map((h: string, idx: number) => (
                <p key={idx} className="mt-0.5">&bull; {h}</p>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Card controls */}
      <div className="flex justify-between items-center gap-3">
        <button 
          onClick={handlePrev}
          className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-slate-300 rounded-lg text-xs border border-slate-800 font-medium transition-colors"
        >
          Trở lại
        </button>

        {!isFlipped && (
          <button 
            onClick={handleCorrect}
            className="px-5 py-2 bg-gradient-to-r from-emerald-600 to-green-600 hover:from-emerald-500 hover:to-green-500 text-white rounded-lg text-xs font-bold transition-all shadow-md shadow-emerald-950/20"
          >
            Tôi đã biết câu trả lời (+10đ)
          </button>
        )}

        <button 
          onClick={handleNext}
          className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-slate-300 rounded-lg text-xs border border-slate-800 font-medium transition-colors"
        >
          Kế tiếp
        </button>
      </div>
    </div>
  );
}

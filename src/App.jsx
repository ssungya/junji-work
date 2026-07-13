import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  ListFilter,
  Plus,
  Search,
  Download,
  AlertTriangle,
  CheckCircle,
  Clock,
  Briefcase,
  Calendar,
  User,
  Settings,
  Trash2,
  Edit,
  Sparkles,
  Copy,
  ChevronDown,
  X,
  FileText,
  UserCheck,
  Building,
  CheckSquare
} from 'lucide-react';
import './App.css';

// --- 초기 모의 데이터 (Mock Data) ---
const INITIAL_TASKS = [
  {
    id: 1,
    requestor: '김민준 대리',
    department: '생산관리팀',
    title: '콘크리트 펌프카 KCP55 유압 실린더 누유 발생 A/S 요청',
    content: '현장 작동 중 KCP55 모델의 3번 붐 유압 실린더에서 작동유 누유가 감지되었습니다. 실린더 씰 교체 및 긴급 점검 수리가 필요합니다. 작업이 지연되면 공기가 늦춰질 우려가 있습니다.',
    category: 'A/S 문의',
    priority: '긴급',
    assignedDept: 'CS/AS팀',
    status: '접수',
    deadline: '2026-07-16',
    createdDate: '2026-07-13',
    summary: 'KCP55 펌프카 유압 실린더 누유에 따른 긴급 정비 요청의 건',
    replyDraft: `안녕하세요, 전진건설로봇 CS/AS팀입니다.

보내주신 "콘크리트 펌프카 KCP55 유압 실린더 누유 발생 A/S 요청" 건에 대해 긴급 접수 완료되었습니다.

본 유압 장비 고장 건은 붐 작동의 안전과 직결되므로 최우선 순위(긴급)로 지정하여 엔지니어를 배정하였습니다. 현장 방문 일정을 확정하기 위해 추가 상세 현장 위치와 담당자 연락처를 회신해 주시기 바랍니다.

[접수 내용 요약]
- 요청 분류: A/S 문의 (우선순위: 긴급)
- 담당 부서: CS/AS팀
- 내용 요약: KCP55 펌프카 유압 실린더 누유에 따른 긴급 정비 요청의 건

조속히 조치해 드리겠습니다. 감사합니다.`
  },
  {
    id: 2,
    requestor: '이성우 과장',
    department: '영업팀',
    title: '사우디아라비아 수출용 장비 견적서 송부 요청',
    content: '사우디 현지 건설 파트너사로부터 콘크리트 펌프카 KCP60 및 KCP38 모델 각 3대에 대한 견적 문의가 인입되었습니다. 특별 사양 단가 조율 및 견적서 작성이 필요합니다.',
    category: '영업 문의',
    priority: '보통',
    assignedDept: '영업팀',
    status: '진행중',
    deadline: '2026-07-20',
    createdDate: '2026-07-13',
    summary: '사우디아라비아 바이어 전달용 KCP60/KCP38 견적 요청의 건',
    replyDraft: `안녕하세요, 전진건설로봇 영업팀입니다.

요청하신 "사우디아라비아 수출용 장비 견적서 송부 요청" 건에 대해 상세 규격 검토 중에 있습니다.

사우디향 수출 전용 사양(모래 필터 추가, 고온 구동 패키지 등) 조율 및 단가 조율을 적용하여 정식 견적서를 발행해 드리겠습니다.

[접수 내용 요약]
- 요청 분류: 영업 문의 (우선순위: 보통)
- 담당 부서: 영업팀
- 내용 요약: 사우디아라비아 바이어 전달용 KCP60/KCP38 견적 요청의 건

작성 완료 후 이메일로 송부드리겠습니다. 감사합니다.`
  },
  {
    id: 3,
    requestor: '박소현 대리',
    department: '품질관리팀',
    title: '신형 실린더 품질 보증 테스트 성적서 발행의 건',
    content: '이번 주 품질 검사가 완료된 신형 유압 실린더 H-900 모델에 대한 품질 보증 테스트 성적서 및 공인 인증 문서 발행을 요청드립니다. 납품 업체 제출용입니다.',
    category: '자료 요청',
    priority: '낮음',
    assignedDept: '기술지원팀',
    status: '완료',
    deadline: '2026-07-15',
    createdDate: '2026-07-12',
    summary: 'H-900 실린더 품질 보증 테스트 성적서 요청의 건',
    replyDraft: `안녕하세요, 전진건설로봇 기술지원팀입니다.

요청하신 "신형 실린더 품질 보증 테스트 성적서 발행의 건"에 대해 조치 완료되었습니다.

첨부 문서로 공인 보증 성적서 PDF 파일과 스캔본을 전송해 드리오니, 파트너사 납품 시 첨부하여 활용하시기 바랍니다.

[접수 내용 요약]
- 요청 분류: 자료 요청 (우선순위: 낮음)
- 담당 부서: 기술지원팀
- 내용 요약: H-900 실린더 품질 보증 테스트 성적서 요청의 건

추가 서류 필요시 연락주십시오. 감사합니다.`
  },
  {
    id: 4,
    requestor: '최재혁 주임',
    department: '구매팀',
    title: '엔진 제어 센서 모듈 해외 발주 긴급 검토',
    content: '유로6 규제 대응용 수입 엔진 제어 센서 모듈의 재고가 임계치 이하로 떨어졌습니다. 독일 Bosch사와의 추가 자재 수입 계약 체결 및 수량 발주를 신속히 진행해야 합니다.',
    category: '부품 문의',
    priority: '긴급',
    assignedDept: '구매팀',
    status: '진행중',
    deadline: '2026-07-14',
    createdDate: '2026-07-13',
    summary: '수입 엔진 제어 센서 모듈 재고 소진에 따른 자재 발주 검토의 건',
    replyDraft: `안녕하세요, 전진건설로봇 구매팀입니다.

요청하신 "엔진 제어 센서 모듈 해외 발주 긴급 검토" 건에 대한 수량 확보 및 계약 검토를 긴급하게 착수했습니다.

독일 측 담당 파트너와 항공 운송 일정을 논의하고 있으며, 단가 조율 및 세관 승인 절차를 간소화하여 재고 차질이 없도록 빠르게 진행하겠습니다.

[접수 내용 요약]
- 요청 분류: 부품 문의 (우선순위: 긴급)
- 담당 부서: 구매팀
- 내용 요약: 수입 엔진 제어 센서 모듈 재고 소진에 따른 자재 발주 검토의 건

진행 상태에 대해 내일 중 다시 안내드리겠습니다. 감사합니다.`
  }
];

// --- 부서 목록 ---
const DEPARTMENTS = [
  '생산관리팀',
  '총무팀',
  '인사팀',
  '영업팀',
  '구매팀',
  '품질관리팀',
  '고객지원팀',
  '일반업무팀'
];

// --- 분류 카테고리 ---
const CATEGORIES = [
  'A/S 문의',
  '납기/출고 문의',
  '영업 문의',
  '부품 문의',
  '자료 요청',
  '품질 이슈',
  '일반 문의'
];

// --- 담당 부서 ---
const ASSIGNED_DEPTS = [
  'CS/AS팀',
  '생산관리팀',
  '영업팀',
  '구매팀',
  '기술지원팀',
  '품질관리팀',
  '총무팀',
  'IT지원팀'
];

export default function App() {
  // --- 상태 관리 ---
  const [tasks, setTasks] = useState(() => {
    const saved = localStorage.getItem('junjin_task_manager_tasks');
    return saved ? JSON.parse(saved) : INITIAL_TASKS;
  });

  const [search, setSearch] = useState('');
  const [filterStatus, setFilterStatus] = useState('all');
  const [filterPriority, setFilterPriority] = useState('all');
  const [filterDept, setFilterDept] = useState('all');
  const [filterCategory, setFilterCategory] = useState('all');
  const [sortBy, setSortBy] = useState('newest');

  // 등록 모달 관련 상태
  const [isRegModalOpen, setIsRegModalOpen] = useState(false);
  const [regRequestor, setRegRequestor] = useState('');
  const [regDept, setRegDept] = useState('생산관리팀');
  const [regTitle, setRegTitle] = useState('');
  const [regContent, setRegContent] = useState('');
  const [regDeadline, setRegDeadline] = useState('');

  // AI 분석 추천 상태
  const [aiAnalyzing, setAiAnalyzing] = useState(false);
  const [aiResult, setAiResult] = useState(null);

  // 등록 폼 내 선택값 (AI 추천값 로드 후 조율용)
  const [selectedCategory, setSelectedCategory] = useState('일반 문의');
  const [selectedPriority, setSelectedPriority] = useState('보통');
  const [selectedAssignedDept, setSelectedAssignedDept] = useState('총무팀');
  const [generatedSummary, setGeneratedSummary] = useState('');
  const [generatedReply, setGeneratedReply] = useState('');

  // 상세 모달 상태
  const [selectedTask, setSelectedTask] = useState(null);

  // 수정 모달 상태
  const [editingTask, setEditingTask] = useState(null);

  // 삭제 확인 모달 상태
  const [deletingTaskId, setDeletingTaskId] = useState(null);

  // 알림 토스트 메시지 상태
  const [toast, setToast] = useState({ show: false, message: '', type: 'success' });

  // 로컬스토리지 저장
  useEffect(() => {
    localStorage.setItem('junjin_task_manager_tasks', JSON.stringify(tasks));
  }, [tasks]);

  // 토스트 메시지 띄우기
  const showToast = (message, type = 'success') => {
    setToast({ show: true, message, type });
    setTimeout(() => {
      setToast({ show: false, message: '', type: 'success' });
    }, 2500);
  };

  // --- 로컬 모의 AI 자동 분석 엔진 ---
  const runMockAI = () => {
    if (!regContent.trim()) {
      showToast('업무 내용을 입력해 주셔야 AI가 분석할 수 있습니다.', 'warning');
      return;
    }

    setAiAnalyzing(true);
    setAiResult(null);

    // AI 가 구동되는 척 연출하기 위해 800ms 딜레이 부여
    setTimeout(() => {
      const text = regContent;
      let category = '일반 문의';
      let priority = '보통';
      let assignedDept = '총무팀';
      let summary = '';
      let replyDraft = '';

      // 키워드 분석
      if (/고장|수리|작동|필터|유압|실린더|소음|누유|부동액|정비|AS|펌프카|센서/.test(text)) {
        category = 'A/S 문의';
        assignedDept = 'CS/AS팀';
        priority = /누유|고장|멈춤|중단|급함|긴급/.test(text) ? '긴급' : '보통';
      } else if (/출고|납기|배송|운송|일정|생산|제조|재고|수출/.test(text)) {
        category = '납기/출고 문의';
        assignedDept = '생산관리팀';
        priority = /긴급|시급|지연|늦어/.test(text) ? '긴급' : '보통';
      } else if (/견적|가격|단가|구매|영업|할인|프로모션|바이어|계약/.test(text)) {
        category = '영업 문의';
        assignedDept = '영업팀';
        priority = '보통';
      } else if (/부품|소모품|발주|자재|수입|체인|기어/.test(text)) {
        category = '부품 문의';
        assignedDept = '구매팀';
        priority = /부족|소진|긴급|지연/.test(text) ? '긴급' : '보통';
      } else if (/도면|성적서|매뉴얼|카탈로그|인증서|도서|자료/.test(text)) {
        category = '자료 요청';
        assignedDept = '기술지원팀';
        priority = '낮음';
      } else if (/불량|클레임|하자|균열|크랙|도장/.test(text)) {
        category = '품질 이슈';
        assignedDept = '품질관리팀';
        priority = '긴급';
      }

      // 한줄 요약 자동 추출/조합
      const firstSentence = text.split(/[.!?\n]/).filter(Boolean)[0] || '';
      summary = `${regTitle ? regTitle.substring(0, 15) : '장비'} 건에 관한 ${category} 분석 및 처리 요청`;
      if (firstSentence) {
        summary = firstSentence.substring(0, 40) + (firstSentence.length > 40 ? '...' : '');
      }

      // 이메일 초안 생성
      replyDraft = `안녕하세요, 전진건설로봇 ${assignedDept}입니다.

보내주신 "${regTitle || '문의 사항'}" 건에 대해 정상적으로 접수되었습니다.

본 요청 사항은 내부 분석 결과 [${category}] 범주에 해당하며, 우선순위는 [${priority}]으로 설정되었습니다. 이에 맞추어 신속하게 대응해 드리겠습니다.

[검토 내용 요약]
- 접수 분류: ${category} (우선순위: ${priority})
- 담당 부서: ${assignedDept}
- 요청 요약: ${summary}

업무 조치가 완료되거나 확인이 필요한 세부 사항이 있는 경우, 즉시 다시 안내해 드리겠습니다.

감사합니다.
전진건설로봇 드림.`;

      // 결과 할당
      const result = { category, priority, assignedDept, summary, replyDraft };
      setAiResult(result);
      
      // 사용자 인터페이스의 폼 값에도 즉시 반영
      setSelectedCategory(category);
      setSelectedPriority(priority);
      setSelectedAssignedDept(assignedDept);
      setGeneratedSummary(summary);
      setGeneratedReply(replyDraft);

      setAiAnalyzing(false);
      showToast('AI 스마트 분석이 완료되었습니다. 추천 값을 검토해 주세요!', 'success');
    }, 900);
  };

  // --- 업무 등록 완료 ---
  const handleRegisterTask = (e) => {
    e.preventDefault();
    if (!regRequestor.trim() || !regTitle.trim() || !regContent.trim() || !regDeadline) {
      showToast('모든 필수 입력 필드를 기입해 주시기 바랍니다.', 'warning');
      return;
    }

    // AI 자동분석을 진행하지 않은 경우 임시 분석 수행
    let finalSummary = generatedSummary;
    let finalReply = generatedReply;

    if (!finalSummary) {
      finalSummary = `${regTitle.substring(0, 20)}... 건의 접수`;
      finalReply = `안녕하세요, 전진건설로봇 ${selectedAssignedDept}입니다.

보내주신 "${regTitle}" 건에 대해 정상 접수되었습니다. 담당자가 내용을 확인한 뒤 조속히 처리해 드리겠습니다.

감사합니다.`;
    }

    const newTask = {
      id: Date.now(),
      requestor: regRequestor,
      department: regDept,
      title: regTitle,
      content: regContent,
      category: selectedCategory,
      priority: selectedPriority,
      assignedDept: selectedAssignedDept,
      status: '접수',
      deadline: regDeadline,
      createdDate: new Date().toISOString().substring(0, 10),
      summary: finalSummary,
      replyDraft: finalReply
    };

    setTasks([newTask, ...tasks]);
    showToast('새 업무 요청이 성공적으로 등록되었습니다.');
    
    // 리셋 및 닫기
    resetRegForm();
  };

  const resetRegForm = () => {
    setRegRequestor('');
    setRegDept('생산관리팀');
    setRegTitle('');
    setRegContent('');
    setRegDeadline('');
    setAiResult(null);
    setSelectedCategory('일반 문의');
    setSelectedPriority('보통');
    setSelectedAssignedDept('총무팀');
    setGeneratedSummary('');
    setGeneratedReply('');
    setIsRegModalOpen(false);
  };

  // --- 업무 정보 수정 완료 ---
  const handleUpdateTask = (e) => {
    e.preventDefault();
    if (!editingTask.requestor.trim() || !editingTask.title.trim() || !editingTask.content.trim() || !editingTask.deadline) {
      showToast('모든 필수 항목을 입력해 주세요.', 'warning');
      return;
    }

    // 수정 완료 처리
    const updated = tasks.map(t => t.id === editingTask.id ? editingTask : t);
    setTasks(updated);
    showToast('업무 정보가 변경되었습니다.');
    setEditingTask(null);
  };

  // --- 업무 삭제 진행 ---
  const triggerDelete = (id) => {
    setDeletingTaskId(id);
  };

  const confirmDelete = () => {
    const updated = tasks.filter(t => t.id !== deletingTaskId);
    setTasks(updated);
    showToast('해당 업무가 삭제되었습니다.', 'info');
    setDeletingTaskId(null);
    if (selectedTask && selectedTask.id === deletingTaskId) {
      setSelectedTask(null);
    }
  };

  // --- 진행 상태 다이렉트 드롭다운 조정 ---
  const handleStatusChange = (taskId, newStatus) => {
    const updated = tasks.map(t => {
      if (t.id === taskId) {
        return { ...t, status: newStatus };
      }
      return t;
    });
    setTasks(updated);
    showToast(`진행 상태가 "${newStatus}"(으)로 변경되었습니다.`, 'success');
  };

  // --- 이메일 초안 복사 ---
  const copyReplyDraft = (text) => {
    navigator.clipboard.writeText(text).then(() => {
      showToast('답변 메일 초안이 클립보드에 복사되었습니다.', 'success');
    }).catch(() => {
      showToast('클립보드 복사에 실패했습니다.', 'warning');
    });
  };

  // --- 대시보드 통계 지표 계산 ---
  const stats = (() => {
    const total = tasks.length;
    const inProgress = tasks.filter(t => t.status === '진행중').length;
    const completed = tasks.filter(t => t.status === '완료').length;
    const urgent = tasks.filter(t => t.priority === '긴급').length;
    
    const todayStr = new Date().toISOString().substring(0, 10);
    const todayRegistered = tasks.filter(t => t.createdDate === todayStr).length;

    // 미완료 상태에서 마감일이 지난 업무 (오늘 날짜 기준)
    const overdue = tasks.filter(t => {
      if (t.status === '완료') return false;
      return t.deadline < todayStr;
    }).length;

    return { total, inProgress, completed, urgent, todayRegistered, overdue };
  })();

  // --- 검색 필터 정렬 가공 엔진 ---
  const processedTasks = (() => {
    let result = [...tasks];

    // 1. 검색어 필터 (제목, 요청자, 요청부서)
    if (search.trim()) {
      const q = search.toLowerCase();
      result = result.filter(
        t =>
          t.title.toLowerCase().includes(q) ||
          t.requestor.toLowerCase().includes(q) ||
          t.department.toLowerCase().includes(q) ||
          (t.summary && t.summary.toLowerCase().includes(q))
      );
    }

    // 2. 상태 필터
    if (filterStatus !== 'all') {
      result = result.filter(t => t.status === filterStatus);
    }

    // 3. 우선순위 필터
    if (filterPriority !== 'all') {
      result = result.filter(t => t.priority === filterPriority);
    }

    // 4. 요청부서 필터
    if (filterDept !== 'all') {
      result = result.filter(t => t.department === filterDept);
    }

    // 5. 분류 카테고리 필터
    if (filterCategory !== 'all') {
      result = result.filter(t => t.category === filterCategory);
    }

    // 6. 정렬 처리
    result.sort((a, b) => {
      if (sortBy === 'newest') return b.id - a.id;
      if (sortBy === 'oldest') return a.id - b.id;
      if (sortBy === 'deadline') return new Date(a.deadline) - new Date(b.deadline);
      if (sortBy === 'priority') {
        const priorityOrder = { '긴급': 1, '보통': 2, '낮음': 3 };
        return (priorityOrder[a.priority] || 9) - (priorityOrder[b.priority] || 9);
      }
      return 0;
    });

    return result;
  })();

  // --- Excel 다운로드 기능 ---
  const downloadExcel = () => {
    if (processedTasks.length === 0) {
      showToast('다운로드할 업무 목록이 비어 있습니다.', 'warning');
      return;
    }

    // 데이터 변환
    const exportData = processedTasks.map((t, idx) => ({
      '번호': idx + 1,
      '신청자': t.requestor,
      '신청부서': t.department,
      '업무제목': t.title,
      '문의유형': t.category,
      '우선순위': t.priority,
      '담당부서': t.assignedDept,
      '진행상태': t.status,
      '마감일': t.deadline,
      '등록일': t.createdDate,
      'AI요약': t.summary
    }));

    // 워크시트 생성
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '업무요청목록');

    // 열 너비 조절
    const maxLen = Math.max(...exportData.map(d => d.업무제목.length));
    worksheet['!cols'] = [
      { wch: 6 },  // 번호
      { wch: 12 }, // 신청자
      { wch: 12 }, // 신청부서
      { wch: Math.min(Math.max(maxLen, 15), 45) }, // 업무제목
      { wch: 12 }, // 문의유형
      { wch: 10 }, // 우선순위
      { wch: 12 }, // 담당부서
      { wch: 10 }, // 진행상태
      { wch: 12 }, // 마감일
      { wch: 12 }, // 등록일
      { wch: 30 }  // AI요약
    ];

    // 저장
    XLSX.writeFile(workbook, `전진건설로봇_업무요청_목록_${new Date().toISOString().substring(0,10)}.xlsx`);
    showToast('Excel 파일 다운로드가 완료되었습니다.', 'success');
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans pb-16">
      {/* --- 탑 네비게이션 --- */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur-md border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-sky-500 flex items-center justify-center text-white shadow-md shadow-blue-200">
                <Sparkles className="w-5 h-5 animate-pulse" />
              </div>
              <div>
                <h1 className="text-lg font-bold text-slate-900 leading-tight">전진건설로봇</h1>
                <p className="text-xs text-slate-500 font-medium">AI 업무요청 · 고객문의 관리 도우미</p>
              </div>
            </div>
            <div className="flex items-center gap-3">
              <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-800 border border-blue-200">
                Vibe Coding 실습용
              </span>
            </div>
          </div>
        </div>
      </header>

      {/* --- 메인 콘텐츠 레이아웃 --- */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8 animate-fade-in">
        
        {/* --- 1. 대시보드 통계 카드 그리드 --- */}
        <section className="grid grid-cols-2 lg:grid-cols-6 gap-4 mb-8">
          {/* 전체 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">총 업무요청</p>
              <h3 className="text-2xl font-extrabold text-slate-800 mt-1">{stats.total}건</h3>
            </div>
            <div className="p-3 bg-slate-50 rounded-xl text-slate-500">
              <Briefcase className="w-5 h-5" />
            </div>
          </div>

          {/* 진행중 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">진행중</p>
              <h3 className="text-2xl font-extrabold text-amber-600 mt-1">{stats.inProgress}건</h3>
            </div>
            <div className="p-3 bg-amber-50 rounded-xl text-amber-600">
              <Clock className="w-5 h-5" />
            </div>
          </div>

          {/* 완료 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">완료</p>
              <h3 className="text-2xl font-extrabold text-emerald-600 mt-1">{stats.completed}건</h3>
            </div>
            <div className="p-3 bg-emerald-50 rounded-xl text-emerald-600">
              <CheckCircle className="w-5 h-5" />
            </div>
          </div>

          {/* 긴급 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">긴급 업무</p>
              <h3 className="text-2xl font-extrabold text-rose-600 mt-1">{stats.urgent}건</h3>
            </div>
            <div className="p-3 bg-rose-50 rounded-xl text-rose-600">
              <AlertTriangle className="w-5 h-5" />
            </div>
          </div>

          {/* 오늘 등록 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">오늘 접수</p>
              <h3 className="text-2xl font-extrabold text-blue-600 mt-1">{stats.todayRegistered}건</h3>
            </div>
            <div className="p-3 bg-blue-50 rounded-xl text-blue-600">
              <Calendar className="w-5 h-5" />
            </div>
          </div>

          {/* 마감 초과 미완료 */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm flex items-center justify-between glow-hover">
            <div>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">마감 초과</p>
              <h3 className="text-2xl font-extrabold text-red-700 mt-1">{stats.overdue}건</h3>
            </div>
            <div className="p-3 bg-red-50 rounded-xl text-red-700">
              <Clock className="w-5 h-5" />
            </div>
          </div>
        </section>

        {/* --- 2. 컨트롤바 (검색, 필터, 등록 버튼) --- */}
        <section className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm mb-6 flex flex-col gap-4">
          <div className="flex flex-col lg:flex-row gap-3 items-stretch lg:items-center">
            {/* 검색바 */}
            <div className="relative flex-1">
              <Search className="w-5 h-5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                placeholder="업무 제목, 요청자, 요청부서 검색..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className="w-full pl-11 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-slate-800 placeholder-slate-400 outline-none focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-100 transition-all font-medium text-sm"
              />
            </div>
            
            {/* 정렬 및 액션 버튼 */}
            <div className="flex flex-wrap gap-2 items-center">
              {/* 정렬 드롭다운 */}
              <div className="relative">
                <select
                  value={sortBy}
                  onChange={(e) => setSortBy(e.target.value)}
                  className="appearance-none bg-slate-50 border border-slate-200 rounded-xl pl-4 pr-10 py-2.5 font-semibold text-xs text-slate-600 outline-none hover:bg-slate-100 cursor-pointer transition-all"
                >
                  <option value="newest">최신 등록순</option>
                  <option value="oldest">오래된 등록순</option>
                  <option value="deadline">마감 기한순</option>
                  <option value="priority">우선순위순</option>
                </select>
                <ChevronDown className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              </div>

              {/* 엑셀 다운로드 */}
              <button
                onClick={downloadExcel}
                className="inline-flex items-center gap-1.5 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-all cursor-pointer border border-slate-200"
              >
                <Download className="w-4 h-4" />
                Excel 다운로드
              </button>

              {/* 신규 등록 버튼 */}
              <button
                onClick={() => setIsRegModalOpen(true)}
                className="inline-flex items-center gap-1.5 px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl shadow-md shadow-blue-150 transition-all cursor-pointer"
              >
                <Plus className="w-4.5 h-4.5" />
                업무 등록
              </button>
            </div>
          </div>

          {/* 상세 필터 행 */}
          <div className="flex flex-wrap gap-2 pt-3 border-t border-slate-100 items-center">
            <div className="inline-flex items-center gap-1.5 text-xs text-slate-400 font-bold mr-1.5">
              <ListFilter className="w-4 h-4" />
              상세 필터
            </div>

            {/* 상태 필터 */}
            <select
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              className="bg-slate-50 hover:bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 cursor-pointer outline-none transition-all"
            >
              <option value="all">진행상태: 전체</option>
              <option value="접수">접수</option>
              <option value="진행중">진행중</option>
              <option value="완료">완료</option>
              <option value="보류">보류</option>
            </select>

            {/* 우선순위 필터 */}
            <select
              value={filterPriority}
              onChange={(e) => setFilterPriority(e.target.value)}
              className="bg-slate-50 hover:bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 cursor-pointer outline-none transition-all"
            >
              <option value="all">우선순위: 전체</option>
              <option value="긴급">긴급</option>
              <option value="보통">보통</option>
              <option value="낮음">낮음</option>
            </select>

            {/* 요청 부서 필터 */}
            <select
              value={filterDept}
              onChange={(e) => setFilterDept(e.target.value)}
              className="bg-slate-50 hover:bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 cursor-pointer outline-none transition-all"
            >
              <option value="all">요청부서: 전체</option>
              {DEPARTMENTS.map(d => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>

            {/* 문의 유형 필터 */}
            <select
              value={filterCategory}
              onChange={(e) => setFilterCategory(e.target.value)}
              className="bg-slate-50 hover:bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 cursor-pointer outline-none transition-all"
            >
              <option value="all">문의유형: 전체</option>
              {CATEGORIES.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>

            {/* 필터 초기화 버튼 */}
            {(filterStatus !== 'all' || filterPriority !== 'all' || filterDept !== 'all' || filterCategory !== 'all' || search) && (
              <button
                onClick={() => {
                  setFilterStatus('all');
                  setFilterPriority('all');
                  setFilterDept('all');
                  setFilterCategory('all');
                  setSearch('');
                }}
                className="text-xs text-blue-600 hover:text-blue-800 font-bold ml-auto hover:underline"
              >
                필터 조건 초기화
              </button>
            )}
          </div>
        </section>

        {/* --- 3. 업무 요청 테이블 목록 --- */}
        <section className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
          <div className="overflow-x-auto custom-scrollbar">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 font-bold text-xs uppercase tracking-wider">
                  <th className="py-4 px-6 text-center w-12">번호</th>
                  <th className="py-4 px-4 w-[38%]">업무 요청 제목</th>
                  <th className="py-4 px-4">요청자 (부서)</th>
                  <th className="py-4 px-4">문의 유형</th>
                  <th className="py-4 px-4">우선순위</th>
                  <th className="py-4 px-4">담당 부서</th>
                  <th className="py-4 px-4">진행 상태</th>
                  <th className="py-4 px-4 text-center">마감일</th>
                  <th className="py-4 px-6 text-center w-28">액션</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-150 text-slate-700 text-sm">
                {processedTasks.length === 0 ? (
                  <tr>
                    <td colSpan="9" className="py-12 text-center text-slate-400 font-semibold bg-slate-50/20">
                      검색 조건과 일치하는 업무 요청 항목이 존재하지 않습니다.
                    </td>
                  </tr>
                ) : (
                  processedTasks.map((t, index) => {
                    const todayStr = new Date().toISOString().substring(0, 10);
                    const isOverdue = t.deadline < todayStr && t.status !== '완료';

                    return (
                      <tr key={t.id} className="hover:bg-slate-50/60 transition-all">
                        {/* 번호 */}
                        <td className="py-4 px-6 text-center font-bold text-slate-400">{index + 1}</td>
                        
                        {/* 업무 제목 */}
                        <td className="py-4 px-4">
                          <button
                            onClick={() => setSelectedTask(t)}
                            className="font-bold text-slate-800 hover:text-blue-600 transition-colors text-left outline-none block max-w-full truncate focus:underline"
                            title="업무 정보 자세히 보기"
                          >
                            {t.title}
                          </button>
                          {t.summary && (
                            <span className="text-xs text-slate-400 font-medium block mt-0.5 truncate max-w-md">
                              AI 요약: {t.summary}
                            </span>
                          )}
                        </td>

                        {/* 요청자 */}
                        <td className="py-4 px-4 font-semibold text-slate-600">
                          {t.requestor}
                          <span className="text-xs font-normal text-slate-400 block mt-0.5">{t.department}</span>
                        </td>

                        {/* 문의 유형 */}
                        <td className="py-4 px-4">
                          <span className="inline-flex px-2 py-1 rounded-md text-xs font-bold bg-slate-100 text-slate-600 border border-slate-200">
                            {t.category}
                          </span>
                        </td>

                        {/* 우선순위 */}
                        <td className="py-4 px-4">
                          <span className={`inline-flex px-2 py-1 rounded-md text-xs font-bold ${
                            t.priority === '긴급' ? 'badge-priority-emergency' :
                            t.priority === '보통' ? 'badge-priority-normal' :
                            'badge-priority-low'
                          }`}>
                            {t.priority}
                          </span>
                        </td>

                        {/* 담당 부서 */}
                        <td className="py-4 px-4 font-bold text-slate-600">{t.assignedDept}</td>

                        {/* 진행상태 셀렉트박스 */}
                        <td className="py-4 px-4">
                          <div className="relative inline-block">
                            <select
                              value={t.status}
                              onChange={(e) => handleStatusChange(t.id, e.target.value)}
                              className={`appearance-none pl-3 pr-8 py-1.5 rounded-lg text-xs font-bold cursor-pointer outline-none border transition-all ${
                                t.status === '접수' ? 'badge-status-received' :
                                t.status === '진행중' ? 'badge-status-inprogress' :
                                t.status === '완료' ? 'badge-status-completed' :
                                'badge-status-onhold'
                              }`}
                            >
                              <option value="접수">접수</option>
                              <option value="진행중">진행중</option>
                              <option value="완료">완료</option>
                              <option value="보류">보류</option>
                            </select>
                            <ChevronDown className="w-3.5 h-3.5 absolute right-2 top-1/2 -translate-y-1/2 pointer-events-none text-slate-500" />
                          </div>
                        </td>

                        {/* 마감일 */}
                        <td className="py-4 px-4 text-center font-bold">
                          <span className={isOverdue ? 'text-rose-600 font-extrabold flex items-center justify-center gap-1' : 'text-slate-600'}>
                            {t.deadline}
                            {isOverdue && <span className="text-[10px] px-1 py-0.5 rounded bg-rose-50 border border-rose-200">초과</span>}
                          </span>
                        </td>

                        {/* 액션 */}
                        <td className="py-4 px-6 text-center">
                          <div className="flex items-center justify-center gap-1">
                            {/* 수정 */}
                            <button
                              onClick={() => setEditingTask({ ...t })}
                              className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer"
                              title="업무 수정"
                            >
                              <Edit className="w-4 h-4" />
                            </button>
                            {/* 삭제 */}
                            <button
                              onClick={() => triggerDelete(t.id)}
                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                              title="업무 삭제"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
          {/* 테이블 푸터 통계 요약 */}
          <div className="bg-slate-50 px-6 py-4 border-t border-slate-200 flex justify-between items-center text-xs font-bold text-slate-500">
            <span>목록 행: {processedTasks.length}건 출력 (전체 {tasks.length}건)</span>
            <span className="text-blue-600">전진건설로봇 AI 업무 분류 엔진 가동중</span>
          </div>
        </section>
      </main>

      {/* ========================================================
          --- 모달창 구현 섹션 ---
      ======================================================== */}

      {/* --- 1. 업무 등록 모달 --- */}
      {isRegModalOpen && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4 animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-3xl shadow-2xl border border-slate-100 animate-scale-in overflow-hidden max-h-[90vh] flex flex-col">
            {/* 헤더 */}
            <div className="px-6 py-5 bg-gradient-to-r from-blue-600 to-sky-500 text-white flex justify-between items-center">
              <div className="flex items-center gap-2">
                <Sparkles className="w-5 h-5" />
                <h3 className="text-base font-extrabold">새로운 업무 요청 등록</h3>
              </div>
              <button onClick={resetRegForm} className="text-white/80 hover:text-white p-1 hover:bg-white/10 rounded-lg transition-all cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* 입력 폼 바디 */}
            <form onSubmit={handleRegisterTask} className="flex-1 overflow-y-auto p-6 space-y-5 custom-scrollbar">
              <div className="grid grid-cols-2 gap-4">
                {/* 신청자 */}
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1.5">신청자명 <span className="text-rose-500">*</span></label>
                  <div className="relative">
                    <User className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      required
                      placeholder="예: 홍길동 대리"
                      value={regRequestor}
                      onChange={(e) => setRegRequestor(e.target.value)}
                      className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-sm font-semibold text-slate-800"
                    />
                  </div>
                </div>

                {/* 신청 부서 */}
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1.5">신청 부서 <span className="text-rose-500">*</span></label>
                  <div className="relative">
                    <Building className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <select
                      value={regDept}
                      onChange={(e) => setRegDept(e.target.value)}
                      className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-sm font-semibold text-slate-700 bg-white"
                    >
                      {DEPARTMENTS.map(d => (
                        <option key={d} value={d}>{d}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              {/* 마감 기한 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1.5">마감 기한 <span className="text-rose-500">*</span></label>
                <div className="relative">
                  <Calendar className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="date"
                    required
                    value={regDeadline}
                    onChange={(e) => setRegDeadline(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-sm font-semibold text-slate-800"
                  />
                </div>
              </div>

              {/* 제목 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1.5">업무 요청 제목 <span className="text-rose-500">*</span></label>
                <div className="relative">
                  <FileText className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    required
                    placeholder="업무 제목을 요약하여 적어주세요."
                    value={regTitle}
                    onChange={(e) => setRegTitle(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-sm font-semibold text-slate-800"
                  />
                </div>
              </div>

              {/* 내용 */}
              <div>
                <div className="flex justify-between items-center mb-1.5">
                  <label className="block text-xs font-bold text-slate-500">업무 상세 내용 <span className="text-rose-500">*</span></label>
                  {/* AI 분석 버튼 */}
                  <button
                    type="button"
                    onClick={runMockAI}
                    disabled={aiAnalyzing}
                    className="inline-flex items-center gap-1.5 px-3 py-1 bg-gradient-to-tr from-violet-600 to-indigo-500 hover:from-violet-700 hover:to-indigo-600 text-white rounded-lg text-xs font-extrabold shadow-sm transition-all cursor-pointer disabled:opacity-50"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    {aiAnalyzing ? 'AI 분석 분석중...' : 'AI 자동 분석 실행'}
                  </button>
                </div>
                <textarea
                  required
                  rows="4"
                  placeholder="장비 결함 상태, 부품 수량, 견적서 납품 등 AI가 세밀한 분류를 처리할 수 있게 상세 내역을 입력해 주세요."
                  value={regContent}
                  onChange={(e) => setRegContent(e.target.value)}
                  className="w-full p-4 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-sm text-slate-800 leading-relaxed resize-none"
                ></textarea>
              </div>

              {/* AI 자동 추천 결과 영역 */}
              {aiResult && (
                <div className="p-5 bg-gradient-to-br from-violet-50/50 to-indigo-50/50 border border-violet-100 rounded-2xl animate-slide-up">
                  <div className="flex items-center gap-2 mb-3">
                    <Sparkles className="w-4.5 h-4.5 text-violet-600" />
                    <span className="text-sm font-extrabold text-violet-800">AI 추천 분석서</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-violet-100 text-violet-700 font-bold ml-1.5">스마트 분류</span>
                  </div>
                  
                  <div className="grid grid-cols-3 gap-3 mb-4">
                    {/* 카테고리 */}
                    <div className="bg-white p-2.5 rounded-xl border border-violet-100">
                      <span className="text-[10px] font-bold text-slate-400 block mb-0.5">추천 문의 유형</span>
                      <select
                        value={selectedCategory}
                        onChange={(e) => setSelectedCategory(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 px-2 py-1 rounded text-xs font-bold text-slate-700 outline-none"
                      >
                        {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    {/* 우선순위 */}
                    <div className="bg-white p-2.5 rounded-xl border border-violet-100">
                      <span className="text-[10px] font-bold text-slate-400 block mb-0.5">추천 우선순위</span>
                      <select
                        value={selectedPriority}
                        onChange={(e) => setSelectedPriority(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 px-2 py-1 rounded text-xs font-bold text-slate-700 outline-none"
                      >
                        <option value="긴급">긴급</option>
                        <option value="보통">보통</option>
                        <option value="낮음">낮음</option>
                      </select>
                    </div>

                    {/* 담당 부서 */}
                    <div className="bg-white p-2.5 rounded-xl border border-violet-100">
                      <span className="text-[10px] font-bold text-slate-400 block mb-0.5">추천 담당 부서</span>
                      <select
                        value={selectedAssignedDept}
                        onChange={(e) => setSelectedAssignedDept(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 px-2 py-1 rounded text-xs font-bold text-slate-700 outline-none"
                      >
                        {ASSIGNED_DEPTS.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                    </div>
                  </div>

                  {/* AI 한줄요약 */}
                  <div className="mb-3">
                    <span className="text-[10px] font-bold text-slate-400 block mb-1">추천 AI 한줄 요약</span>
                    <input
                      type="text"
                      value={generatedSummary}
                      onChange={(e) => setGeneratedSummary(e.target.value)}
                      className="w-full bg-white border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-bold text-slate-700 outline-none"
                    />
                  </div>

                  {/* 이메일 초안 */}
                  <div>
                    <div className="flex justify-between items-center mb-1">
                      <span className="text-[10px] font-bold text-slate-400">추천 자동 답변 메일 초안</span>
                      <button
                        type="button"
                        onClick={() => copyReplyDraft(generatedReply)}
                        className="text-[10px] text-violet-600 hover:text-violet-800 font-bold flex items-center gap-0.5"
                      >
                        <Copy className="w-3 h-3" />
                        초안 복사
                      </button>
                    </div>
                    <textarea
                      rows="3"
                      value={generatedReply}
                      onChange={(e) => setGeneratedReply(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-lg text-xs text-slate-600 leading-relaxed outline-none resize-none font-mono"
                    ></textarea>
                  </div>
                </div>
              )}

              {/* 하단 푸터 버튼 */}
              <div className="pt-4 border-t border-slate-100 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={resetRegForm}
                  className="px-4 py-2 border border-slate-200 text-slate-500 font-bold text-xs rounded-xl hover:bg-slate-50 transition-all cursor-pointer"
                >
                  취소
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl shadow-md transition-all cursor-pointer"
                >
                  등록 완료
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* --- 2. 업무 상세 보기 모달 --- */}
      {selectedTask && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4 animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-2xl shadow-2xl border border-slate-100 animate-scale-in overflow-hidden max-h-[85vh] flex flex-col">
            {/* 헤더 */}
            <div className="px-6 py-5 border-b border-slate-150 flex justify-between items-start bg-slate-50">
              <div className="space-y-1">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className={`inline-flex px-2 py-0.5 rounded text-[10px] font-bold ${
                    selectedTask.priority === '긴급' ? 'badge-priority-emergency' :
                    selectedTask.priority === '보통' ? 'badge-priority-normal' :
                    'badge-priority-low'
                  }`}>
                    {selectedTask.priority}
                  </span>
                  <span className="inline-flex px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">
                    {selectedTask.category}
                  </span>
                  <span className={`inline-flex px-2 py-0.5 rounded text-[10px] font-bold ${
                    selectedTask.status === '접수' ? 'badge-status-received' :
                    selectedTask.status === '진행중' ? 'badge-status-inprogress' :
                    selectedTask.status === '완료' ? 'badge-status-completed' :
                    'badge-status-onhold'
                  }`}>
                    {selectedTask.status}
                  </span>
                </div>
                <h3 className="text-base font-extrabold text-slate-800 leading-snug">{selectedTask.title}</h3>
              </div>
              <button
                onClick={() => setSelectedTask(null)}
                className="text-slate-400 hover:text-slate-600 p-1 hover:bg-slate-100 rounded-lg transition-all cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* 본문 바디 */}
            <div className="p-6 overflow-y-auto flex-1 space-y-6 custom-scrollbar text-sm text-slate-700 leading-relaxed">
              {/* 작성자 정보 */}
              <div className="grid grid-cols-2 gap-4 bg-slate-50 p-4 rounded-xl border border-slate-150">
                <div className="flex items-center gap-2">
                  <UserCheck className="w-4 h-4 text-slate-400" />
                  <span className="text-slate-500 font-bold">요청자:</span>
                  <span className="font-bold text-slate-700">{selectedTask.requestor} ({selectedTask.department})</span>
                </div>
                <div className="flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-slate-400" />
                  <span className="text-slate-500 font-bold">마감기한:</span>
                  <span className="font-bold text-rose-600">{selectedTask.deadline}</span>
                </div>
                <div className="flex items-center gap-2">
                  <Building className="w-4 h-4 text-slate-400" />
                  <span className="text-slate-500 font-bold">담당부서:</span>
                  <span className="font-extrabold text-blue-600">{selectedTask.assignedDept}</span>
                </div>
                <div className="flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-slate-400" />
                  <span className="text-slate-500 font-bold">등록일:</span>
                  <span className="font-bold text-slate-600">{selectedTask.createdDate}</span>
                </div>
              </div>

              {/* 요청 상세 내용 */}
              <div>
                <span className="block text-xs font-bold text-slate-400 mb-1.5 uppercase tracking-wide">요청 상세 내용</span>
                <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-150 whitespace-pre-line text-slate-800 leading-loose min-h-[100px]">
                  {selectedTask.content}
                </div>
              </div>

              {/* AI 요약 */}
              {selectedTask.summary && (
                <div>
                  <span className="inline-flex items-center gap-1 text-xs font-bold text-violet-600 mb-1.5">
                    <Sparkles className="w-3.5 h-3.5" />
                    AI 업무 요약
                  </span>
                  <div className="bg-violet-50/40 border border-violet-100/80 p-4 rounded-xl text-slate-700 font-bold">
                    {selectedTask.summary}
                  </div>
                </div>
              )}

              {/* 이메일 답변 초안 */}
              {selectedTask.replyDraft && (
                <div>
                  <div className="flex justify-between items-center mb-1.5">
                    <span className="inline-flex items-center gap-1 text-xs font-bold text-violet-600">
                      <Sparkles className="w-3.5 h-3.5" />
                      AI 자동 메일 답변 초안
                    </span>
                    <button
                      onClick={() => copyReplyDraft(selectedTask.replyDraft)}
                      className="inline-flex items-center gap-1 text-xs text-blue-600 hover:text-blue-800 font-extrabold"
                    >
                      <Copy className="w-3.5 h-3.5" />
                      이메일 초안 복사
                    </button>
                  </div>
                  <pre className="bg-slate-900 text-slate-200 p-4 rounded-xl text-xs font-mono whitespace-pre-wrap leading-relaxed overflow-x-auto">
                    {selectedTask.replyDraft}
                  </pre>
                </div>
              )}
            </div>

            {/* 하단 버튼 */}
            <div className="px-6 py-4 border-t border-slate-150 flex justify-end gap-2 bg-slate-50">
              <button
                onClick={() => {
                  setEditingTask({ ...selectedTask });
                  setSelectedTask(null);
                }}
                className="px-4 py-2 border border-slate-200 bg-white hover:bg-slate-50 text-slate-600 font-bold text-xs rounded-xl cursor-pointer transition-all"
              >
                수정
              </button>
              <button
                onClick={() => setSelectedTask(null)}
                className="px-5 py-2 bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs rounded-xl cursor-pointer transition-all"
              >
                확인
              </button>
            </div>
          </div>
        </div>
      )}

      {/* --- 3. 업무 수정 모달 --- */}
      {editingTask && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4 animate-fade-in">
          <div className="bg-white rounded-3xl w-full max-w-2xl shadow-2xl border border-slate-100 animate-scale-in overflow-hidden max-h-[85vh] flex flex-col">
            <div className="px-6 py-5 bg-slate-800 text-white flex justify-between items-center">
              <h3 className="text-base font-extrabold">업무 요청 정보 수정</h3>
              <button onClick={() => setEditingTask(null)} className="text-white/80 hover:text-white p-1 hover:bg-white/10 rounded-lg cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleUpdateTask} className="p-6 overflow-y-auto flex-1 space-y-4 custom-scrollbar text-sm">
              <div className="grid grid-cols-2 gap-4">
                {/* 신청자 */}
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">신청자</label>
                  <input
                    type="text"
                    required
                    value={editingTask.requestor}
                    onChange={(e) => setEditingTask({ ...editingTask, requestor: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 font-semibold"
                  />
                </div>
                {/* 신청부서 */}
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">신청부서</label>
                  <select
                    value={editingTask.department}
                    onChange={(e) => setEditingTask({ ...editingTask, department: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 font-semibold bg-white"
                  >
                    {DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
              </div>

              {/* 마감일 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1">마감 기한</label>
                <input
                  type="date"
                  required
                  value={editingTask.deadline}
                  onChange={(e) => setEditingTask({ ...editingTask, deadline: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 font-semibold"
                />
              </div>

              {/* 제목 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1">업무 제목</label>
                <input
                  type="text"
                  required
                  value={editingTask.title}
                  onChange={(e) => setEditingTask({ ...editingTask, title: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 font-semibold"
                />
              </div>

              {/* 내용 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1">상세 내용</label>
                <textarea
                  required
                  rows="3"
                  value={editingTask.content}
                  onChange={(e) => setEditingTask({ ...editingTask, content: e.target.value })}
                  className="w-full p-3 border border-slate-200 rounded-xl outline-none focus:border-blue-500 leading-relaxed resize-none"
                ></textarea>
              </div>

              {/* 분류 정보 조율 */}
              <div className="grid grid-cols-3 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-150">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 mb-1">문의 유형</label>
                  <select
                    value={editingTask.category}
                    onChange={(e) => setEditingTask({ ...editingTask, category: e.target.value })}
                    className="w-full bg-white border border-slate-200 px-2 py-1.5 rounded-lg text-xs font-bold text-slate-700 outline-none"
                  >
                    {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 mb-1">우선순위</label>
                  <select
                    value={editingTask.priority}
                    onChange={(e) => setEditingTask({ ...editingTask, priority: e.target.value })}
                    className="w-full bg-white border border-slate-200 px-2 py-1.5 rounded-lg text-xs font-bold text-slate-700 outline-none"
                  >
                    <option value="긴급">긴급</option>
                    <option value="보통">보통</option>
                    <option value="낮음">낮음</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 mb-1">담당 부서</label>
                  <select
                    value={editingTask.assignedDept}
                    onChange={(e) => setEditingTask({ ...editingTask, assignedDept: e.target.value })}
                    className="w-full bg-white border border-slate-200 px-2 py-1.5 rounded-lg text-xs font-bold text-slate-700 outline-none"
                  >
                    {ASSIGNED_DEPTS.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
              </div>

              {/* 한줄요약 수동 수정 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1">AI 요약 문구 수정</label>
                <input
                  type="text"
                  value={editingTask.summary}
                  onChange={(e) => setEditingTask({ ...editingTask, summary: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl outline-none focus:border-blue-500 font-semibold"
                />
              </div>

              {/* 초안 메일 수정 */}
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-1">답변 메일 초안 수정</label>
                <textarea
                  rows="3"
                  value={editingTask.replyDraft}
                  onChange={(e) => setEditingTask({ ...editingTask, replyDraft: e.target.value })}
                  className="w-full p-3 border border-slate-200 rounded-xl outline-none focus:border-blue-500 text-xs leading-relaxed resize-none font-mono"
                ></textarea>
              </div>

              <div className="pt-4 border-t border-slate-100 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setEditingTask(null)}
                  className="px-4 py-2 border border-slate-200 text-slate-500 font-bold text-xs rounded-xl hover:bg-slate-50 cursor-pointer"
                >
                  취소
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs rounded-xl cursor-pointer"
                >
                  저장
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* --- 4. 삭제 확인 커스텀 경고 모달 --- */}
      {deletingTaskId && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4 animate-fade-in">
          <div className="bg-white rounded-2xl w-full max-w-sm p-6 shadow-2xl border border-slate-150 animate-scale-in text-center">
            <div className="w-12 h-12 bg-rose-50 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-4 border border-rose-200">
              <AlertTriangle className="w-6 h-6" />
            </div>
            <h4 className="text-base font-extrabold text-slate-800">정말로 삭제하시겠습니까?</h4>
            <p className="text-xs text-slate-400 mt-2 font-medium">삭제된 업무 요청 데이터는 다시 복구할 수 없습니다.</p>
            <div className="mt-6 flex justify-center gap-2">
              <button
                onClick={() => setDeletingTaskId(null)}
                className="px-4 py-2 border border-slate-200 text-slate-500 font-bold text-xs rounded-xl hover:bg-slate-50 cursor-pointer"
              >
                취소
              </button>
              <button
                onClick={confirmDelete}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-xl cursor-pointer shadow-sm"
              >
                삭제 진행
              </button>
            </div>
          </div>
        </div>
      )}

      {/* --- 5. 글로벌 알림 토스트 팝업 --- */}
      {toast.show && (
        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-2 px-4 py-3 rounded-xl shadow-lg border animate-slide-up text-xs font-bold bg-slate-900 text-white border-slate-800">
          {toast.type === 'success' && <CheckCircle className="w-4 h-4 text-emerald-400" />}
          {toast.type === 'warning' && <AlertTriangle className="w-4 h-4 text-amber-400" />}
          {toast.type === 'info' && <Clock className="w-4 h-4 text-sky-400" />}
          <span>{toast.message}</span>
        </div>
      )}
    </div>
  );
}

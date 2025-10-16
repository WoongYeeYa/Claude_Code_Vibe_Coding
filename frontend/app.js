// Word to HWP 변환 시스템 - 프론트엔드 JavaScript

// 전역 상태
let currentFileId = null;
let currentFile = null;

// DOM 요소
const elements = {
    // 상태 표시
    statusIndicator: document.getElementById('statusIndicator'),

    // 탭
    tabBtns: document.querySelectorAll('.tab-btn'),
    convertTab: document.getElementById('convert-tab'),
    templatesTab: document.getElementById('templates-tab'),

    // 파일 업로드
    uploadArea: document.getElementById('uploadArea'),
    fileInput: document.getElementById('fileInput'),
    uploadedFile: document.getElementById('uploadedFile'),
    fileName: document.getElementById('fileName'),
    fileSize: document.getElementById('fileSize'),
    removeFileBtn: document.getElementById('removeFileBtn'),

    // 액션 버튼
    convertBtn: document.getElementById('convertBtn'),
    saveTemplateBtn: document.getElementById('saveTemplateBtn'),

    // 결과
    resultSection: document.getElementById('resultSection'),
    resultMessage: document.getElementById('resultMessage'),
    downloadBtn: document.getElementById('downloadBtn'),

    // 템플릿
    templatesList: document.getElementById('templatesList'),
    refreshTemplatesBtn: document.getElementById('refreshTemplatesBtn'),

    // 로딩
    loadingOverlay: document.getElementById('loadingOverlay'),
    loadingText: document.getElementById('loadingText'),

    // 모달
    templateModal: document.getElementById('templateModal'),
    templateName: document.getElementById('templateName'),
    closeModalBtn: document.getElementById('closeModalBtn'),
    cancelModalBtn: document.getElementById('cancelModalBtn'),
    confirmSaveBtn: document.getElementById('confirmSaveBtn')
};

// 초기화
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
    setupEventListeners();
    checkSystemStatus();
});

// 앱 초기화
function initializeApp() {
    console.log('앱 초기화...');
}

// 이벤트 리스너 설정
function setupEventListeners() {
    // 탭 전환
    elements.tabBtns.forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    // 파일 업로드
    elements.uploadArea.addEventListener('click', () => elements.fileInput.click());
    elements.fileInput.addEventListener('change', handleFileSelect);

    // 드래그 앤 드롭
    elements.uploadArea.addEventListener('dragover', handleDragOver);
    elements.uploadArea.addEventListener('dragleave', handleDragLeave);
    elements.uploadArea.addEventListener('drop', handleDrop);

    // 파일 제거
    elements.removeFileBtn.addEventListener('click', removeFile);

    // 변환 버튼
    elements.convertBtn.addEventListener('click', convertToHWP);

    // 템플릿 저장 버튼
    elements.saveTemplateBtn.addEventListener('click', openTemplateModal);

    // 다운로드 버튼
    elements.downloadBtn.addEventListener('click', downloadHWP);

    // 템플릿 새로고침
    elements.refreshTemplatesBtn.addEventListener('click', loadTemplates);

    // 모달
    elements.closeModalBtn.addEventListener('click', closeTemplateModal);
    elements.cancelModalBtn.addEventListener('click', closeTemplateModal);
    elements.confirmSaveBtn.addEventListener('click', saveTemplate);

    // 모달 외부 클릭 시 닫기
    elements.templateModal.addEventListener('click', (e) => {
        if (e.target === elements.templateModal) {
            closeTemplateModal();
        }
    });
}

// 시스템 상태 확인
async function checkSystemStatus() {
    try {
        const response = await fetch('/api/health');
        const data = await response.json();

        if (data.status === 'ok') {
            if (data.hwp_installed) {
                elements.statusIndicator.classList.add('online');
                elements.statusIndicator.querySelector('.status-text').textContent =
                    `시스템 정상 (한글 ${data.version || '설치됨'})`;
            } else {
                elements.statusIndicator.classList.add('error');
                elements.statusIndicator.querySelector('.status-text').textContent =
                    '한컴오피스 한글 미설치';
            }
        }
    } catch (error) {
        console.error('상태 확인 실패:', error);
        elements.statusIndicator.classList.add('error');
        elements.statusIndicator.querySelector('.status-text').textContent = '서버 연결 실패';
    }
}

// 탭 전환
function switchTab(tabName) {
    // 버튼 활성화
    elements.tabBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabName);
    });

    // 탭 내용 표시
    if (tabName === 'convert') {
        elements.convertTab.classList.add('active');
        elements.templatesTab.classList.remove('active');
    } else if (tabName === 'templates') {
        elements.convertTab.classList.remove('active');
        elements.templatesTab.classList.add('active');
        loadTemplates();
    }
}

// 파일 선택
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        displayFile(file);
    }
}

// 드래그 오버
function handleDragOver(e) {
    e.preventDefault();
    elements.uploadArea.classList.add('dragover');
}

// 드래그 리브
function handleDragLeave(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');
}

// 드롭
function handleDrop(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('dragover');

    const file = e.dataTransfer.files[0];
    if (file) {
        displayFile(file);
    }
}

// 파일 표시
function displayFile(file) {
    // 파일 확장자 확인
    const validExtensions = ['.docx', '.doc'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        showNotification('DOCX 또는 DOC 파일만 업로드 가능합니다', 'error');
        return;
    }

    currentFile = file;

    // UI 업데이트
    elements.uploadArea.style.display = 'none';
    elements.uploadedFile.style.display = 'flex';
    elements.fileName.textContent = file.name;
    elements.fileSize.textContent = formatFileSize(file.size);

    // 결과 섹션 숨기기
    elements.resultSection.style.display = 'none';

    // 버튼 활성화
    elements.convertBtn.disabled = false;
    elements.saveTemplateBtn.disabled = false;

    // 파일 업로드
    uploadFile(file);
}

// 파일 제거
function removeFile() {
    currentFile = null;
    currentFileId = null;

    elements.uploadArea.style.display = 'block';
    elements.uploadedFile.style.display = 'none';
    elements.fileInput.value = '';

    elements.convertBtn.disabled = true;
    elements.saveTemplateBtn.disabled = true;

    elements.resultSection.style.display = 'none';
}

// 파일 업로드
async function uploadFile(file) {
    try {
        showLoading('파일 업로드 중...');

        const formData = new FormData();
        formData.append('file', file);

        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            currentFileId = data.file_id;
            console.log('파일 업로드 성공:', data);
        } else {
            throw new Error(data.message || '업로드 실패');
        }
    } catch (error) {
        console.error('파일 업로드 실패:', error);
        showNotification('파일 업로드에 실패했습니다: ' + error.message, 'error');
        removeFile();
    } finally {
        hideLoading();
    }
}

// HWP 변환
async function convertToHWP() {
    if (!currentFileId) {
        showNotification('파일을 먼저 업로드해주세요', 'error');
        return;
    }

    try {
        showLoading('HWP로 변환 중...');

        const formData = new FormData();
        formData.append('file_id', currentFileId);

        const response = await fetch('/api/convert', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        // 응답 상태 코드 확인
        if (!response.ok) {
            throw new Error(data.detail || `서버 오류 (${response.status})`);
        }

        if (data.success) {
            console.log('변환 성공:', data);
            showResult('HWP 파일이 생성되었습니다');
        } else {
            throw new Error(data.detail || data.message || '변환 실패');
        }
    } catch (error) {
        console.error('변환 실패:', error);
        showNotification('HWP 변환에 실패했습니다: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// HWP 다운로드
async function downloadHWP() {
    if (!currentFileId) {
        showNotification('변환된 파일이 없습니다', 'error');
        return;
    }

    try {
        const url = `/api/download/${currentFileId}`;
        const link = document.createElement('a');
        link.href = url;
        link.download = `converted_${currentFileId}.hwp`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showNotification('다운로드를 시작합니다', 'success');
    } catch (error) {
        console.error('다운로드 실패:', error);
        showNotification('다운로드에 실패했습니다', 'error');
    }
}

// 템플릿 모달 열기
function openTemplateModal() {
    if (!currentFile) {
        showNotification('파일을 먼저 업로드해주세요', 'error');
        return;
    }

    elements.templateModal.style.display = 'flex';
    elements.templateName.value = '';
    elements.templateName.focus();
}

// 템플릿 모달 닫기
function closeTemplateModal() {
    elements.templateModal.style.display = 'none';
}

// 템플릿 저장
async function saveTemplate() {
    const templateName = elements.templateName.value.trim();

    if (!templateName) {
        showNotification('템플릿 이름을 입력해주세요', 'error');
        return;
    }

    if (!currentFile) {
        showNotification('파일을 먼저 업로드해주세요', 'error');
        return;
    }

    try {
        showLoading('템플릿 저장 중...');
        closeTemplateModal();

        const formData = new FormData();
        formData.append('file', currentFile);
        formData.append('template_name', templateName);

        const response = await fetch('/api/template/save', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            showNotification('템플릿이 저장되었습니다', 'success');
            console.log('템플릿 저장 성공:', data);
        } else {
            throw new Error(data.detail || '저장 실패');
        }
    } catch (error) {
        console.error('템플릿 저장 실패:', error);
        showNotification('템플릿 저장에 실패했습니다: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// 템플릿 목록 로드
async function loadTemplates() {
    try {
        const response = await fetch('/api/templates');
        const data = await response.json();

        if (data.success) {
            displayTemplates(data.templates);
        } else {
            throw new Error('템플릿 로드 실패');
        }
    } catch (error) {
        console.error('템플릿 로드 실패:', error);
        showNotification('템플릿 목록을 불러올 수 없습니다', 'error');
    }
}

// 템플릿 표시
function displayTemplates(templates) {
    if (templates.length === 0) {
        elements.templatesList.innerHTML = `
            <div class="empty-state">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path>
                    <polyline points="13 2 13 9 20 9"></polyline>
                </svg>
                <p>저장된 템플릿이 없습니다</p>
                <p class="hint">파일 변환 탭에서 템플릿을 저장해보세요</p>
            </div>
        `;
        return;
    }

    elements.templatesList.innerHTML = templates.map(template => `
        <div class="template-card">
            <div class="template-header">
                <svg class="template-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path>
                    <polyline points="13 2 13 9 20 9"></polyline>
                </svg>
                <div class="template-info">
                    <p class="template-name" title="${template.template_name}">${template.template_name}</p>
                    <p class="template-filename" title="${template.original_filename}">${template.original_filename}</p>
                </div>
            </div>
            <p class="template-date">${formatDate(template.created_at)}</p>
            <div class="template-actions">
                <button class="btn-use" onclick="useTemplate('${template.template_id}')">사용</button>
                <button class="btn-delete" onclick="deleteTemplate('${template.template_id}')">삭제</button>
            </div>
        </div>
    `).join('');
}

// 템플릿 사용
async function useTemplate(templateId) {
    try {
        showLoading('템플릿 사용 중...');

        const response = await fetch(`/api/template/use/${templateId}`, {
            method: 'POST'
        });

        const data = await response.json();

        if (data.success) {
            currentFileId = data.file_id;
            showNotification('템플릿으로 HWP 파일이 생성되었습니다', 'success');

            // 변환 탭으로 이동하고 결과 표시
            switchTab('convert');
            showResult('템플릿을 사용하여 HWP 파일이 생성되었습니다');
        } else {
            throw new Error(data.detail || '템플릿 사용 실패');
        }
    } catch (error) {
        console.error('템플릿 사용 실패:', error);
        showNotification('템플릿 사용에 실패했습니다: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// 템플릿 삭제
async function deleteTemplate(templateId) {
    if (!confirm('이 템플릿을 삭제하시겠습니까?')) {
        return;
    }

    try {
        showLoading('템플릿 삭제 중...');

        const response = await fetch(`/api/template/${templateId}`, {
            method: 'DELETE'
        });

        const data = await response.json();

        if (data.success) {
            showNotification('템플릿이 삭제되었습니다', 'success');
            loadTemplates();
        } else {
            throw new Error(data.detail || '삭제 실패');
        }
    } catch (error) {
        console.error('템플릿 삭제 실패:', error);
        showNotification('템플릿 삭제에 실패했습니다: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// 결과 표시
function showResult(message) {
    elements.resultMessage.textContent = message;
    elements.resultSection.style.display = 'block';
    elements.resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// 로딩 표시
function showLoading(message = '처리 중...') {
    elements.loadingText.textContent = message;
    elements.loadingOverlay.style.display = 'flex';
}

// 로딩 숨기기
function hideLoading() {
    elements.loadingOverlay.style.display = 'none';
}

// 알림 표시
function showNotification(message, type = 'info') {
    // 간단한 alert 사용 (나중에 토스트 알림으로 개선 가능)
    alert(message);
    console.log(`[${type.toUpperCase()}] ${message}`);
}

// 파일 크기 포맷
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// 날짜 포맷
function formatDate(dateString) {
    const date = new Date(dateString);
    const now = new Date();
    const diff = now - date;
    const diffDays = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (diffDays === 0) {
        return '오늘';
    } else if (diffDays === 1) {
        return '어제';
    } else if (diffDays < 7) {
        return `${diffDays}일 전`;
    } else {
        return date.toLocaleDateString('ko-KR');
    }
}

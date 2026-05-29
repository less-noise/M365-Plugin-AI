/**
 * Word AI 助手 - 主逻辑
 */

// 全局变量
let selectedText = '';
let aiResult = '';

// System Prompts
const SYSTEM_PROMPTS = {
  polish: `你是一位专业的学术论文编辑，精通中文学术写作规范。你的任务是润色用户提供的文字，具体要求：
1. 保持原文的核心观点和逻辑结构不变
2. 提升表达的学术性和严谨性，避免口语化表达
3. 修正语法错误和不通顺的句子
4. 优化段落连贯性
5. 直接输出润色后的完整文字，不要添加任何解释或对比说明`,

  format: `你是一位专业的文档排版助手。你的任务是整理用户提供文字的格式，具体要求：
1. 统一标点符号用法（中文语境使用全角标点）
2. 修正段落缩进和空行问题（段首空两格）
3. 规范数字与单位的书写（如"5个"改为"5 个"，中英文之间加空格）
4. 直接输出整理后的完整文字，不要添加任何解释`,

  tone: `你是一位写作助手。请将用户提供的文字调整为更{tone}的语气和表达风格。保持原意不变，直接输出调整后的完整文字，不要添加任何解释。`,

  summary: `你是一位专业的学术摘要写作专家。请根据用户提供的文字，生成一段规范的学术摘要，要求：
1. 摘要长度约150-300字
2. 包含研究目的、方法、结果和结论四个要素
3. 使用第三人称，语言客观严谨
4. 直接输出摘要内容，不加标题`,

  extend: `你是一位专业的学术论文写作助手。请根据用户提供的文字，在保持相同学术风格的前提下进行续写扩展。要求：
1. 续写内容与原文逻辑连贯
2. 扩展后总长度约为原文的1.5-2倍
3. 保持学术论文的写作规范
4. 直接输出扩展后的完整文字（包含原文）`
};

// Office 初始化
Office.onReady(function(info) {
  if (info.host === Office.HostType.Word) {
    console.log('Word AI 助手已加载');
    initializeUI();
  }
});

/**
 * 初始化 UI 事件绑定
 */
function initializeUI() {
  document.getElementById('getSelectionBtn').addEventListener('click', getSelectedText);
  document.getElementById('processBtn').addEventListener('click', processText);
  document.getElementById('replaceBtn').addEventListener('click', replaceOriginalText);
  document.getElementById('copyBtn').addEventListener('click', copyResult);
}

/**
 * 获取文档中选中的文字
 */
async function getSelectedText() {
  hideError();

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();

      selectedText = selection.text;

      const previewBox = document.getElementById('previewBox');
      if (selectedText && selectedText.trim()) {
        previewBox.textContent = selectedText;
        previewBox.style.color = '#333';
      } else {
        previewBox.textContent = '请先在 Word 文档中选中要处理的文字';
        previewBox.style.color = '#999';
        selectedText = '';
      }
    });
  } catch (error) {
    showError('无法读取文档内容，请确认插件权限');
    console.error(error);
  }
}

/**
 * 开始处理文字
 */
async function processText() {
  hideError();

  if (!selectedText || !selectedText.trim()) {
    showError('请先在文档中选中要处理的文字');
    return;
  }

  const functionRadio = document.querySelector('input[name="function"]:checked');
  const functionType = functionRadio.value;
  const additionalInput = document.getElementById('additionalInput').value.trim();

  // 构建 System Prompt
  let systemPrompt = SYSTEM_PROMPTS[functionType];
  if (functionType === 'tone' && additionalInput) {
    systemPrompt = systemPrompt.replace('{tone}', additionalInput);
  } else if (functionType === 'tone') {
    systemPrompt = systemPrompt.replace('{tone}', '正式学术');
  }

  showLoading(true);
  disableButtons(true);

  const resultBox = document.getElementById('resultBox');
  resultBox.textContent = '';
  aiResult = '';

  try {
    const model = await getModelForApp('word');

    aiResult = await callDeepSeek(systemPrompt, selectedText, model, (chunk) => {
      resultBox.textContent += chunk;
      resultBox.scrollTop = resultBox.scrollHeight;
    });

    document.getElementById('replaceBtn').disabled = false;
    document.getElementById('copyBtn').disabled = false;

  } catch (error) {
    showError(error.message);
    console.error(error);
  } finally {
    showLoading(false);
    disableButtons(false);
  }
}

/**
 * 替换原文
 * 注意：Word.run 创建的 context 在 run 完成后即被释放，
 * 不能跨 context 使用对象引用，因此这里重新获取当前选区进行替换。
 */
async function replaceOriginalText() {
  hideError();

  if (!aiResult || !aiResult.trim()) {
    showError('没有可替换的内容');
    return;
  }

  try {
    await Word.run(async (context) => {
      // 在当前的 Word.run 上下文中获取选区（不能复用旧 context 的对象）
      const selection = context.document.getSelection();
      selection.insertText(aiResult, Word.InsertLocation.replace);
      await context.sync();
    });

    showSuccess('已成功替换原文');

  } catch (error) {
    showError('无法替换文档内容，请重新选中文字后再试');
    console.error(error);
  }
}

/**
 * 复制结果到剪贴板
 */
async function copyResult() {
  if (!aiResult || !aiResult.trim()) {
    return;
  }

  try {
    await navigator.clipboard.writeText(aiResult);
    showCopySuccess();
  } catch (error) {
    // 备用方案：使用传统方法
    const textarea = document.createElement('textarea');
    textarea.value = aiResult;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    showCopySuccess();
  }
}

function showCopySuccess() {
  const copySuccess = document.getElementById('copySuccess');
  copySuccess.classList.add('show');
  setTimeout(() => {
    copySuccess.classList.remove('show');
  }, 2000);
}

function showLoading(show) {
  document.getElementById('loadingDiv').style.display = show ? 'flex' : 'none';
}

function disableButtons(disabled) {
  document.getElementById('getSelectionBtn').disabled = disabled;
  document.getElementById('processBtn').disabled = disabled;
}

function showError(message) {
  const errorBox = document.getElementById('errorBox');
  errorBox.textContent = '❌ ' + message;
  errorBox.style.display = 'block';
}

function showSuccess(message) {
  const errorBox = document.getElementById('errorBox');
  errorBox.textContent = '✅ ' + message;
  errorBox.className = 'success-box';
  errorBox.style.display = 'block';
  setTimeout(() => {
    hideError();
  }, 3000);
}

function hideError() {
  const errorBox = document.getElementById('errorBox');
  errorBox.style.display = 'none';
  errorBox.className = 'error-box';
}

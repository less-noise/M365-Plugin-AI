/**
 * Excel AI 助手 - 主逻辑
 */

// 全局变量
let selectedData = null;
let aiResult = '';

// Placeholder 示例
const PLACEHOLDER_HINTS = {
  clean: '示例：去除重复行，将日期列统一格式为YYYY-MM-DD',
  question: '示例：各月份销售额分别是多少？哪个月最高？',
  trend: '示例：请描述这份数据的整体趋势和异常点',
  formula: '示例：计算B列中大于1000的数值之和'
};

// System Prompts
const SYSTEM_PROMPTS = {
  clean: `你是一位数据处理专家。用户会提供一份表格数据（JSON格式）和清洗要求。请分析数据后，给出具体的处理建议和操作步骤，以清晰的中文列表形式输出。如果可以生成 Excel 公式辅助处理，请一并提供。不要直接修改数据，给出可操作的指导即可。`,
  
  question: `你是一位数据分析助手。用户会提供一份表格数据（JSON格式）和问题。请基于数据回答问题，给出准确的数字和分析结论，使用清晰的中文表达。`,
  
  trend: `你是一位数据分析师。用户会提供一份表格数据（JSON格式）。请分析数据的整体趋势、关键特征和异常值，以专业但易懂的中文段落描述，适合直接用于报告。`,
  
  formula: `你是一位 Excel 专家。用户会描述他想要的计算需求，以及相关数据的列信息。请给出对应的 Excel 公式，并附上一句话解释公式的含义。直接给出公式，格式为：公式：=XXX，说明：XXX。`
};

// Office 初始化
Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    console.log('Excel AI 助手已加载');
    initializeUI();
  }
});

/**
 * 初始化 UI 事件绑定
 */
function initializeUI() {
  // 功能选择变化时更新 placeholder
  const functionRadios = document.querySelectorAll('input[name="function"]');
  functionRadios.forEach(radio => {
    radio.addEventListener('change', updatePlaceholder);
  });
  
  // 开始处理按钮
  document.getElementById('processBtn').addEventListener('click', processData);
  
  // 复制按钮
  document.getElementById('copyBtn').addEventListener('click', copyResult);
  
  // 初始化 placeholder
  updatePlaceholder();
}

/**
 * 更新 placeholder 提示
 */
function updatePlaceholder() {
  const functionRadio = document.querySelector('input[name="function"]:checked');
  const functionType = functionRadio.value;
  const placeholderHint = document.getElementById('placeholderHint');
  const userInput = document.getElementById('userInput');
  
  placeholderHint.textContent = PLACEHOLDER_HINTS[functionType];
  userInput.placeholder = '请输入你的问题或指令...';
}

/**
 * 读取选区数据
 */
async function readSelectedRange() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      
      // 数据量限制：超过 500 行截取前 500 行
      let values = range.values;
      let rowCount = range.rowCount;
      let truncated = false;
      
      if (rowCount > 500) {
        values = values.slice(0, 500);
        rowCount = 500;
        truncated = true;
      }
      
      // 计算单元格数量
      const cellCount = rowCount * range.columnCount;
      
      // 更新选区信息显示
      const selectionInfo = document.getElementById('selectionInfo');
      selectionInfo.textContent = `已选择: ${range.address} (${cellCount}个单元格)`;
      
      if (truncated) {
        selectionInfo.textContent += ' (数据较多，仅分析前500行)';
      }
      
      // 保存数据
      selectedData = {
        values: values,
        address: range.address,
        rowCount: rowCount,
        columnCount: range.columnCount,
        truncated: truncated
      };
    });
    
    return true;
  } catch (error) {
    showError('无法读取选区数据，请确认插件权限');
    console.error(error);
    return false;
  }
}

/**
 * 开始处理数据
 */
async function processData() {
  hideError();
  
  // 获取用户输入
  const userInput = document.getElementById('userInput').value.trim();
  
  if (!userInput) {
    showError('请输入问题或指令后再点击处理');
    return;
  }
  
  // 读取选区数据
  const success = await readSelectedRange();
  if (!success) {
    return;
  }
  
  // 获取选中的功能
  const functionRadio = document.querySelector('input[name="function"]:checked');
  const functionType = functionRadio.value;
  
  // 构建 System Prompt
  const systemPrompt = SYSTEM_PROMPTS[functionType];
  
  // 构建用户消息（包含数据）
  let userMessage = userInput;
  
  // 对于需要数据的功能，将数据转为 JSON 并附加
  if (functionType !== 'formula' && selectedData) {
    const dataJson = JSON.stringify(selectedData.values, null, 2);
    userMessage = `数据（JSON格式）：\n${dataJson}\n\n问题/要求：${userInput}`;
  }
  
  // 显示加载状态
  showLoading(true);
  disableButtons(true);
  
  // 清空结果框
  const resultBox = document.getElementById('resultBox');
  resultBox.textContent = '';
  aiResult = '';
  
  try {
    // 获取模型
    const model = await getModelForApp('excel');
    
    // 调用 DeepSeek API（流式输出）
    aiResult = await callDeepSeek(systemPrompt, userMessage, model, (chunk) => {
      resultBox.textContent += chunk;
      // 自动滚动到底部
      resultBox.scrollTop = resultBox.scrollHeight;
    });
    
    // 启用复制按钮
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
 * 复制结果到剪贴板
 */
async function copyResult() {
  if (!aiResult || !aiResult.trim()) {
    return;
  }
  
  try {
    // 使用现代 Clipboard API
    await navigator.clipboard.writeText(aiResult);
    
    // 显示复制成功提示
    const copySuccess = document.getElementById('copySuccess');
    copySuccess.classList.add('show');
    
    setTimeout(() => {
      copySuccess.classList.remove('show');
    }, 2000);
    
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
    
    // 显示复制成功提示
    const copySuccess = document.getElementById('copySuccess');
    copySuccess.classList.add('show');
    
    setTimeout(() => {
      copySuccess.classList.remove('show');
    }, 2000);
  }
}

/**
 * 显示/隐藏加载状态
 */
function showLoading(show) {
  const loadingDiv = document.getElementById('loadingDiv');
  loadingDiv.style.display = show ? 'flex' : 'none';
}

/**
 * 禁用/启用按钮
 */
function disableButtons(disabled) {
  document.getElementById('processBtn').disabled = disabled;
}

/**
 * 显示错误提示
 */
function showError(message) {
  const errorBox = document.getElementById('errorBox');
  errorBox.textContent = '❌ ' + message;
  errorBox.style.display = 'block';
}

/**
 * 隐藏错误提示
 */
function hideError() {
  const errorBox = document.getElementById('errorBox');
  errorBox.style.display = 'none';
}
/**
 * PowerPoint AI 助手 - 主逻辑
 */

// 全局变量
let aiResult = '';

// Placeholder 示例
const PLACEHOLDER_HINTS = {
  outline: '示例：人工智能在医疗领域的应用现状与前景',
  content: '示例：粘贴你的大纲或描述每页主题，AI 将为每页生成具体文字',
  notes: '示例：请简要描述这张幻灯片的主题，或直接点击生成（将读取当前幻灯片标题）'
};

// System Prompts
const SYSTEM_PROMPTS = {
  outline: `你是一位专业的演讲策划专家。请根据用户提供的主题，生成一份结构清晰的演讲大纲，要求：
1. 包含5-8张幻灯片的标题
2. 每张幻灯片下列出3-5个要点（一级列表）
3. 大纲结构：封面 → 目录 → 正文若干页 → 总结 → 致谢/Q&A
4. 使用 Markdown 格式输出（## 为幻灯片标题，- 为要点）`,
  
  content: `你是一位专业的幻灯片内容撰写专家。请根据用户提供的大纲或主题描述，为每张幻灯片生成具体的文字内容，要求：
1. 每页标题清晰
2. 正文使用简洁的要点式表达（每点不超过20字）
3. 语言专业、精炼，适合幻灯片展示
4. 使用 Markdown 格式输出`,
  
  notes: `你是一位演讲教练。请根据用户提供的幻灯片标题或内容，生成一段适合演讲者朗读的备注稿，要求：
1. 长度约100-200字
2. 语言自然流畅，适合口头表达
3. 包含对幻灯片内容的补充说明
4. 直接输出备注文字，不加标题`
};

// Office 初始化
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    console.log('PPT AI 助手已加载');
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
  
  // 开始生成按钮
  document.getElementById('processBtn').addEventListener('click', processContent);
  
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
  
  // 根据功能类型设置不同的 placeholder
  if (functionType === 'outline') {
    userInput.placeholder = '请输入演讲或报告主题...';
  } else if (functionType === 'content') {
    userInput.placeholder = '粘贴你的大纲或描述每页主题...';
  } else if (functionType === 'notes') {
    userInput.placeholder = '请简要描述这张幻灯片的主题...';
  }
}

/**
 * 读取当前幻灯片标题
 */
async function getCurrentSlideTitle() {
  try {
    await PowerPoint.run(async (context) => {
      // 获取当前演示文稿
      const presentation = context.presentation;
      
      // 尝试获取当前选中的幻灯片
      // 注意：PowerPoint API 可能有限制，这里尝试读取第一个幻灯片作为示例
      const slides = presentation.slides;
      slides.load('items');
      await context.sync();
      
      if (slides.items.length > 0) {
        // 获取第一个幻灯片（或尝试获取当前幻灯片）
        const slide = slides.items[0];
        slide.load('shapes');
        await context.sync();
        
        // 遍历 shapes 找到 title placeholder
        for (const shape of slide.shapes.items) {
          shape.load(['placeholderType', 'textFrame']);
          await context.sync();
          
          if (shape.placeholderType === 'Title' || shape.placeholderType === 'CenterTitle') {
            shape.textFrame.load('text');
            await context.sync();
            
            if (shape.textFrame.text) {
              return shape.textFrame.text;
            }
          }
        }
      }
      
      return null;
    });
  } catch (error) {
    console.warn('无法读取幻灯片标题:', error);
    return null;
  }
}

/**
 * 开始生成内容
 */
async function processContent() {
  hideError();
  
  // 获取选中的功能
  const functionRadio = document.querySelector('input[name="function"]:checked');
  const functionType = functionRadio.value;
  
  // 获取用户输入
  let userInput = document.getElementById('userInput').value.trim();
  
  // 对于生成备注功能，如果没有用户输入，尝试读取幻灯片标题
  if (functionType === 'notes' && !userInput) {
    const slideTitle = await getCurrentSlideTitle();
    if (slideTitle) {
      userInput = slideTitle;
      // 更新输入框显示
      document.getElementById('userInput').value = slideTitle;
    } else {
      showError('请输入幻灯片主题或描述');
      return;
    }
  }
  
  // 检查输入
  if (!userInput) {
    showError('请输入主题或内容后再点击生成');
    return;
  }
  
  // 构建 System Prompt
  const systemPrompt = SYSTEM_PROMPTS[functionType];
  
  // 显示加载状态
  showLoading(true);
  disableButtons(true);
  
  // 清空结果框
  const resultBox = document.getElementById('resultBox');
  resultBox.textContent = '';
  aiResult = '';
  
  try {
    // 获取模型（PPT 使用 deepseek-v4-pro）
    const model = await getModelForApp('ppt');
    
    // 调用 DeepSeek API（流式输出）
    aiResult = await callDeepSeek(systemPrompt, userInput, model, (chunk) => {
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
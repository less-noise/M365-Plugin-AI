/**
 * DeepSeek API 统一调用模块
 * 使用 Anthropic 兼容端点进行 API 调用
 */

/**
 * 调用 DeepSeek AI
 * @param {string} systemPrompt - 系统提示词
 * @param {string} userMessage - 用户消息内容
 * @param {string} model - 模型名称（来自 config）
 * @param {function} onChunk - 流式输出回调（每次收到文本片段时调用）
 * @returns {Promise<string>} - 完整响应文本
 */
async function callDeepSeek(systemPrompt, userMessage, model, onChunk) {
  // 读取配置
  const config = await loadConfig();
  const apiKey = config.apiKey;
  
  // 检查 API Key
  if (!apiKey || apiKey === 'placeholder' || apiKey === '在此填入你的 DeepSeek API Key') {
    throw new Error('请先在 config/config.json 中填写你的 DeepSeek API Key');
  }
  
  try {
    const response = await fetch('https://api.deepseek.com/anthropic/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: model,
        max_tokens: 4096,
        stream: true,
        system: systemPrompt,
        messages: [
          { role: 'user', content: userMessage }
        ]
      })
    });
    
    // 处理错误响应
    if (!response.ok) {
      if (response.status === 401) {
        throw new Error('API Key 无效，请检查 config.json 中的密钥是否正确');
      } else if (response.status === 429) {
        throw new Error('请求过于频繁，请稍后重试');
      } else if (response.status === 402) {
        throw new Error('API 余额不足，请检查账户余额');
      } else {
        throw new Error(`API 请求失败: ${response.status}`);
      }
    }
    
    // 流式读取响应
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullText = '';
    
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      
      const chunk = decoder.decode(value, { stream: true });
      
      // 解析 SSE 格式：data: {"type":"content_block_delta","delta":{"text":"..."}}
      const lines = chunk.split('\n').filter(l => l.startsWith('data: '));
      
      for (const line of lines) {
        try {
          const dataStr = line.slice(6).trim();
          if (dataStr === '[DONE]') continue;
          
          const data = JSON.parse(dataStr);
          if (data.type === 'content_block_delta' && data.delta?.text) {
            const textChunk = data.delta.text;
            fullText += textChunk;
            if (onChunk) {
              onChunk(textChunk);
            }
          }
        } catch (parseError) {
          // 忽略解析错误，继续处理下一行
          console.warn('解析 SSE 数据时出错:', parseError);
        }
      }
    }
    
    return fullText;
    
  } catch (error) {
    if (error.name === 'TypeError' && error.message.includes('fetch')) {
      throw new Error('网络连接失败，请检查网络后重试');
    }
    throw error;
  }
}

/**
 * 非流式调用 DeepSeek AI（用于简单场景）
 * @param {string} systemPrompt - 系统提示词
 * @param {string} userMessage - 用户消息内容
 * @param {string} model - 模型名称
 * @returns {Promise<string>} - 完整响应文本
 */
async function callDeepSeekSimple(systemPrompt, userMessage, model) {
  const config = await loadConfig();
  const apiKey = config.apiKey;
  
  if (!apiKey || apiKey === 'placeholder' || apiKey === '在此填入你的 DeepSeek API Key') {
    throw new Error('请先在 config/config.json 中填写你的 DeepSeek API Key');
  }
  
  try {
    const response = await fetch('https://api.deepseek.com/anthropic/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: model,
        max_tokens: 4096,
        stream: false,
        system: systemPrompt,
        messages: [
          { role: 'user', content: userMessage }
        ]
      })
    });
    
    if (!response.ok) {
      if (response.status === 401) {
        throw new Error('API Key 无效，请检查 config.json 中的密钥是否正确');
      } else if (response.status === 429) {
        throw new Error('请求过于频繁，请稍后重试');
      } else if (response.status === 402) {
        throw new Error('API 余额不足，请检查账户余额');
      } else {
        throw new Error(`API 请求失败: ${response.status}`);
      }
    }
    
    const data = await response.json();
    return data.content[0].text;
    
  } catch (error) {
    if (error.name === 'TypeError' && error.message.includes('fetch')) {
      throw new Error('网络连接失败，请检查网络后重试');
    }
    throw error;
  }
}
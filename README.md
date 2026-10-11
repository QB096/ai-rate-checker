# ai-rate-checker — 小说「AI 率」检测工具

按章节切分小说 txt/md，调用 DeepSeek 逐块判别「AI 生成概率」，输出 Markdown 报告 + Excel 明细。

## 使用

```bash
pip install httpx openpyxl
python ai_rate_checker.py 小说.txt          # CLI
python ai_rate_checker.py --help            # 参数说明
```

首次使用编辑 `config.json`，填入你的 DeepSeek API Key：

```json
{ "base_url": "https://api.deepseek.com", "api_key": "YOUR_DEEPSEEK_API_KEY", ... }
```

- `测试样例.txt` — 内置测试文本
- `concurrency` 控制并发块数，`threshold` 为报告中的「可疑」阈值

> ⚠️ config.json 含密钥，请勿提交真实值。

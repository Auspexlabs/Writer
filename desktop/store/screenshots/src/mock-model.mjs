// A scripted stand-in for the model service (the OpenAI-compatible /chat/completions stream the engine speaks to a
// "custom" provider), so the AI screenshot shows a real assistant turn without calling a model: the engine runs the
// writer commands below on the sample file, and the editor diffs, marks and lists the change itself.
import { createServer } from 'node:http';

const F = '市场部周报.docx';
export const PROMPT = '把本周进展改成要点、数字加粗，再起草下周计划';
const SCENE = {
  commands: [
    `set ${F} /body/paragraph[2] --prop list=bullet --prop md="海报已在 **120 家**门店上墙"`,
    `add ${F} /body --type paragraph --after /body/paragraph[2] --prop list=bullet --prop md="种草笔记 **36 篇**，互动 **8.6 万次**"`,
    `add ${F} /body --type paragraph --after /body/paragraph[3] --prop list=bullet --prop md="短信覆盖 **42 万**人，打开率 **12%**"`,
    `add ${F} /body --type table --prop style=TableGrid --prop width=100% --prop data='[["日期","事项","负责"],["9 月 28 日","补齐华南 18 家门店晚到的首批物料","运营组"],["9 月 30 日","发布首发倒计时海报，上线小程序预点单","市场部"],["10 月 1 日","新品首发，每两小时汇总一次各区销量","数据组"]]'`,
  ],
  reply: '已把「本周进展」拆成三条要点、关键数字加粗；「下周计划」按上市节奏起草了三项，你看看要不要调整。',
};

const chunk = (res, delta, finish = null) => res.write('data: ' + JSON.stringify({ object: 'chat.completion.chunk', choices: [{ index: 0, delta, finish_reason: finish }] }) + '\n\n');

/** Starts the stand-in on a free port; resolves { url, close }. */
export function startMock() {
  const server = createServer((req, res) => {
    let body = '';
    req.on('data', d => { body += d; });
    req.on('end', () => {
      const messages = (JSON.parse(body || '{}').messages) || [];
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      if (messages.some(m => m.role === 'tool')) {           // the commands ran: answer in words
        chunk(res, { role: 'assistant', content: SCENE.reply }, null);
        chunk(res, {}, 'stop');
      } else {                                                 // first request: edit the file
        chunk(res, { role: 'assistant', tool_calls: SCENE.commands.map((command, index) => ({ index, id: 'call_' + index, type: 'function', function: { name: 'writer', arguments: JSON.stringify({ command }) } })) });
        chunk(res, {}, 'tool_calls');
      }
      res.end('data: [DONE]\n\n');
    });
  });
  return new Promise(resolve => server.listen(0, '127.0.0.1', () => resolve({ url: `http://127.0.0.1:${server.address().port}/v1`, close: () => server.close() })));
}

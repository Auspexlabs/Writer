// A scripted stand-in for the model service (the OpenAI-compatible /chat/completions stream the engine speaks to a
// "custom" provider), so the AI screenshots show a real assistant turn without calling a model: the engine runs the
// writer commands below on the sample file, and the editor diffs, marks and lists the change itself.
// The scene is picked by the request the user typed (shots.mjs types SCENES[*].prompt).
import { createServer } from 'node:http';

export const SCENES = {
  // the hero: the plan's dates move a week on (the opening line, the four rows of the schedule, one note)
  dates: {
    file: '咖啡节活动方案.docx', prompt: '活动改到下个周末了，把文中的日期都改过来',
    commands: [
      `set 咖啡节活动方案.docx /body/paragraph[1] --prop text="本方案用于十月第三个周末在滨江公园举办的城市咖啡节，涵盖场地、日程、分工和预算，供筹备组讨论后定稿。"`,
      `set 咖啡节活动方案.docx /body/table[1]/row[2]/cell[1] --prop text="10 月 17 日 09:00"`,
      `set 咖啡节活动方案.docx /body/table[1]/row[3]/cell[1] --prop text="10 月 17 日 14:00"`,
      `set 咖啡节活动方案.docx /body/table[1]/row[4]/cell[1] --prop text="10 月 18 日 10:00"`,
      `set 咖啡节活动方案.docx /body/table[1]/row[5]/cell[1] --prop text="10 月 18 日 17:00"`,
      `set 咖啡节活动方案.docx /body/paragraph[4] --prop text="所有摊位需在 10 月 16 日 18:00 前完成布置。"`,
    ],
    reply: '已把日期整体顺延一周：开头一处、日程表四处、注意事项一处。',
  },
  // the AI section: the budget gets a 备注 column for the two items under 50%
  budget: {
    file: '咖啡节预算.xlsx', prompt: '进度低于 50% 的项目，在旁边备注「需跟进」',
    commands: [
      `set 咖啡节预算.xlsx /sheet[1]/cell[F1] --prop value=备注 --prop bold=true --prop fill=E8F0EB`,
      `set 咖啡节预算.xlsx /sheet[1]/cell[F5] --prop value=需跟进`,
      `set 咖啡节预算.xlsx /sheet[1]/cell[F7] --prop value=需跟进`,
    ],
    reply: '进度低于 50% 的有两项：工作坊耗材（45%）和安保与保洁（40%）。已加上「备注」列，这两行标注「需跟进」。',
  },
};

const chunk = (res, delta, finish = null) => res.write('data: ' + JSON.stringify({ object: 'chat.completion.chunk', choices: [{ index: 0, delta, finish_reason: finish }] }) + '\n\n');

/** Starts the stand-in on a free port; resolves { url, close }. */
export function startMock() {
  const server = createServer((req, res) => {
    let body = '';
    req.on('data', d => { body += d; });
    req.on('end', () => {
      const messages = (JSON.parse(body || '{}').messages) || [];
      const text = messages.filter(m => m.role === 'user').map(m => typeof m.content === 'string' ? m.content : JSON.stringify(m.content)).join('\n');
      const scene = Object.values(SCENES).find(s => text.includes(s.prompt)) || SCENES.dates;
      res.writeHead(200, { 'Content-Type': 'text/event-stream' });
      if (messages.some(m => m.role === 'tool')) {           // the commands ran: answer in words
        chunk(res, { role: 'assistant', content: scene.reply }, null);
        chunk(res, {}, 'stop');
      } else {                                                 // first request: edit the file
        chunk(res, { role: 'assistant', tool_calls: scene.commands.map((command, index) => ({ index, id: 'call_' + index, type: 'function', function: { name: 'writer', arguments: JSON.stringify({ command }) } })) });
        chunk(res, {}, 'tool_calls');
      }
      res.end('data: [DONE]\n\n');
    });
  });
  return new Promise(resolve => server.listen(0, '127.0.0.1', () => resolve({ url: `http://127.0.0.1:${server.address().port}/v1`, close: () => server.close() })));
}

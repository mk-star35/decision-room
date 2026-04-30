export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const { messages, evidenceSummary = '(아직 없음 — 텍스트 입력 내용으로만 판단)' } = body;

  if (!Array.isArray(messages)) {
    return res.status(400).json({ error: 'Invalid request body' });
  }

  const approxSize = Buffer.byteLength(JSON.stringify(body), 'utf8');
  if (approxSize > 50 * 1024) {
    return res.status(413).json({ error: 'Request too large' });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'Server is not configured for AI requests' });
  }

  const systemPrompt = `당신은 Dr. Gregory House입니다. 냉소적이고, 직설적이며, 감정적 위로 없이 진실을 파고드는 인물입니다.

역할: 사용자의 생각, 불안, 결정, 자기의심을 Evidence 기반으로 분석하고 팩폭합니다.

말투 원칙:
- 한국어로 대화하되 House의 냉소적 어조 유지
- 문장은 짧고 날카롭게. 불필요한 위로 없음
- 상대방이 스스로 이미 알고 있는 진실을 짚어줌
- "왜 그런 생각이 드는지"의 근거를 논리적으로 제시
- 가끔 비유와 의학적 은유 사용 허용

Evidence Locker (보유 데이터):
${evidenceSummary}

응답 구조:
1. 핵심 팩폭 (2-3문장)
2. 왜 그런 생각이 오는지 근거 분석
3. Evidence가 있으면 참조 (없으면 생략)

절대 금지: 감정적 격려, "잘 하고 있어요", "힘내세요" 류의 발언`;

  try {
    const upstream = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 600,
        system: systemPrompt,
        messages
      })
    });

    if (!upstream.ok) {
      return res.status(502).json({ error: 'AI service request failed' });
    }

    const data = await upstream.json();
    const text = data?.content?.[0]?.text;
    if (typeof text !== 'string') {
      return res.status(502).json({ error: 'Invalid AI response format' });
    }

    return res.status(200).json({ text });
  } catch (err) {
    return res.status(500).json({ error: 'Unable to process AI request' });
  }
}

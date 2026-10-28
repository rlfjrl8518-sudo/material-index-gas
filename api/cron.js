// Vercel Cron(매일 08시 KST) → 아침 요약 발송. Vercel이 CRON_SECRET을 Bearer로 실어 보낸다.
const { accessToken, appUrl, runGas } = require('../lib/web');

module.exports = async (req, res) => {
  if (!process.env.CRON_SECRET || req.headers.authorization !== 'Bearer ' + process.env.CRON_SECRET) return res.status(401).end();
  try {
    const { result } = await runGas({ fn: 'sendDailyDigest', args: [], token: await accessToken(), appUrl: appUrl(req), user: 'cron' });
    res.status(200).json(result);
  } catch (e) {
    console.error('digest 실패', e);
    res.status(500).json({ error: e.message });
  }
};

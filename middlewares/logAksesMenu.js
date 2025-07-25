module.exports = async function logAksesMenu(req, res, next) {
  const path = req.path;
  const user = req.user;

  if (!user || path.startsWith('/static') || path === '/favicon.ico') {
    return next();
  }

  const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
  const userAgent = req.headers['user-agent'];

  try {
    await db.execute(
      `INSERT INTO log_akses_menu (user_id, username, menu_path, ip_address, user_agent)
       VALUES (?, ?, ?, ?, ?)`,
      [user.id, user.username, path, ipAddress, userAgent]
    );
  } catch (err) {
    console.error('❌ Gagal log akses menu:', err);
  }

  next();
};

module.exports = function checkAkses(allowed) {
  return function (req, res, next) {
    try {
      let aksesUser = req.user?.akses || [];

      // Jika bentuknya string, parse ke array
      if (typeof aksesUser === 'string') {
        aksesUser = JSON.parse(aksesUser);
      }

      const isAllowed = allowed.some(code => aksesUser.includes(code));
      if (!isAllowed) {
        return res.status(403).send('❌ Akses ditolak');
      }

      next();
    } catch (err) {
      console.error('Error parsing akses user:', err);
      return res.status(403).send('❌ Akses tidak valid');
    }
  };
};

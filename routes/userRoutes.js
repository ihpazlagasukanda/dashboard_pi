const express = require('express');
const router = express.Router();
const userController = require('../controllers/userController');


// Middleware untuk pengecekan level admin
// function requireLevel(level) {
//     return function (req, res, next) {
//         if (!req.user || req.user.level < level) {
//             return res.status(403).render('access-denied'); // Menampilkan halaman access-denied
//         }
//         next();
//     };
// }

// Tampilkan halaman manajemen user
router.get('/manage-users', userController.getAllUser);

// Tambah user/user baru
router.post('/manage-users', userController.createUser);

// Update level atau kabupaten user
router.put('/manage-users/:id', userController.updateUser);

router.put('/manage-users/:id/akses', userController.updateAksesUser); // ini artinya /admin/manage-users/:id/akses


router.get('/manage-users/:id', userController.getUserById);

// Hapus user/user
router.delete('/manage-users/:id', userController.deleteUser);

module.exports = router;

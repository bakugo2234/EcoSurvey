// Admin controller: Handles user management, Excel import, account approval, system stats, and survey grading.
const { Op } = require('sequelize');
const { User, Survey, SurveyResponse, Participation, Notification, FAQ } = require('../models');
const emailService = require('../services/emailService');
const logger = require('../utils/logger');
const { getPagination } = require('../utils/pagination');
const ExcelJS = require('exceljs');
const bcrypt = require('bcrypt');

// Retrieves paginated user list with role, status, and keyword search filters.
exports.getUsers = async (req, res) => {
  try {
    const { role, status, search } = req.query;
    const where = {};
    if (role)   where.role   = role;
    if (status) where.status = status;
    if (search) {
      where[Op.or] = [
        { full_name:        { [Op.like]: `%${search}%` } },
        { username:         { [Op.like]: `%${search}%` } },
        { email:            { [Op.like]: `%${search}%` } },
        { student_staff_id: { [Op.like]: `%${search}%` } },
      ];
    }
    
    const { page, limit, offset } = getPagination(req.query);
    const { count, rows } = await User.findAndCountAll({
      where,
      attributes: { exclude: ['password_hash'] },
      order: [['created_at', 'DESC']],
      limit,
      offset,
    });

    res.json({
      users: rows,
      total: count,
      page,
      totalPages: Math.ceil(count / limit),
    });
  } catch (err) {
    logger.error('getUsers error:', err);
    res.status(500).json({ message: 'Failed to fetch users.' });
  }
};

// Generates and downloads a clean sample Excel template (.xlsx) for bulk user import.
exports.downloadUserTemplate = async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('User Import Template');

    worksheet.columns = [
      { header: 'Họ và tên (*)', key: 'full_name', width: 25 },
      { header: 'Tên đăng nhập (*)', key: 'username', width: 20 },
      { header: 'Email (*)', key: 'email', width: 30 },
      { header: 'Mật khẩu (*)', key: 'password', width: 20 },
      { header: 'Vai trò (Student/Staff) (*)', key: 'role', width: 25 },
      { header: 'Mã SV / Cán bộ (*)', key: 'student_staff_id', width: 22 },
      { header: 'Lớp (Bắt buộc với Student)', key: 'class_name', width: 25 },
      { header: 'Khoa / Phòng ban (*)', key: 'department', width: 30 },
    ];
    
     worksheet.addRow({
      full_name: 'Nguyễn Văn A',
      username: 'nguyenvana',
      email: 'nguyenvana@ecosurvey.edu.vn',
      password: 'Password123',
      role: 'Student',
      student_staff_id: 'SV2026001',
      class_name: 'K20-CNTT1',
      department: 'Công nghệ thông tin',
    });
    worksheet.addRow({
      full_name: 'Trần Thị B',
      username: 'tranthib',
      email: 'tranthib@ecosurvey.edu.vn',
      password: 'Password123',
      role: 'Staff',
      student_staff_id: 'CB2026001',
      class_name: '',
      department: 'Phòng Đào tạo',
    });
    
    // Styling the header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1B4D3E' },
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 28;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="EcoSurvey_User_Import_Template.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    logger.error('downloadUserTemplate error:', err);
    res.status(500).json({ message: 'Không thể tạo file mẫu Excel.' });
  }
};

// Imports user accounts in bulk from an Excel file (.xlsx).
exports.importUsers = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Vui lòng đính kèm file Excel (.xlsx).' });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return res.status(400).json({ message: 'File Excel không có dữ liệu (Worksheet rỗng).' });
    }

    let successful = 0;
    let failed = 0;
    const errors = [];
    const seenUsernames = new Set();
    const seenEmails = new Set();

    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      if (!row.hasValues) continue;

      const full_name = row.getCell(1).text?.trim();
      const username = row.getCell(2).text?.trim();
      const email = row.getCell(3).text?.trim()?.toLowerCase();
      const password = row.getCell(4).text?.trim();
      let role = row.getCell(5).text?.trim();
      const student_staff_id = row.getCell(6).text?.trim() || null;
      const class_name = row.getCell(7).text?.trim() || null;
      const department = row.getCell(8).text?.trim() || null;

      // 1. Kiểm tra các trường cơ bản bắt buộc
      if (!full_name || !username || !email || !password) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Thiếu thông tin cơ bản bắt buộc (Họ tên, Tên đăng nhập, Email, Mật khẩu).`);
        continue;
      }

      // 2. Chặn tuyệt đối không cho phép import tài khoản Admin qua Excel
      if (role === 'Admin') {
        failed++;
        errors.push(`Dòng ${rowNumber}: Không được phép import tài khoản Quản trị viên (Admin). Chỉ chấp nhận 'Student' hoặc 'Staff'.`);
        continue;
      }

      // 3. Kiểm tra vai trò hợp lệ (chỉ chấp nhận Student hoặc Staff)
      if (!role || !['Student', 'Staff'].includes(role)) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Vai trò '${role || ''}' không hợp lệ. Chỉ chấp nhận 'Student' hoặc 'Staff'.`);
        continue;
      }

      // 4. Kiểm tra thông tin bắt buộc theo từng vai trò để tránh tài khoản thiếu dữ liệu truy cập hệ thống
      if (role === 'Student') {
        if (!student_staff_id || !class_name || !department) {
          failed++;
          errors.push(`Dòng ${rowNumber} (Sinh viên): Bắt buộc phải có đầy đủ Mã SV, Lớp và Khoa/Phòng ban.`);
          continue;
        }
      } else if (role === 'Staff') {
        if (!student_staff_id || !department) {
          failed++;
          errors.push(`Dòng ${rowNumber} (Cán bộ): Bắt buộc phải có đầy đủ Mã cán bộ và Khoa/Phòng ban.`);
          continue;
        }
      }

      // 5. Kiểm tra định dạng Email
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(email)) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Email '${email}' không đúng định dạng.`);
        continue;
      }

      // 6. Kiểm tra độ dài mật khẩu tối thiểu
      if (password.length < 6) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Mật khẩu cho tài khoản '${username}' phải có ít nhất 6 ký tự.`);
        continue;
      }

      // 7. Kiểm tra trùng lặp trong chính file Excel
      if (seenUsernames.has(username.toLowerCase())) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Tên đăng nhập '${username}' bị trùng lặp trong file.`);
        continue;
      }
      if (seenEmails.has(email)) {
        failed++;
        errors.push(`Dòng ${rowNumber}: Email '${email}' bị trùng lặp trong file.`);
        continue;
      }
      seenUsernames.add(username.toLowerCase());
      seenEmails.add(email);

      // 8. Kiểm tra trùng lặp trong Database và tạo tài khoản
      try {
        const existing = await User.findOne({
          where: { [Op.or]: [{ username }, { email }] },
          attributes: ['id', 'username', 'email']
        });
        if (existing) {
          failed++;
          if (existing.username.toLowerCase() === username.toLowerCase()) {
            errors.push(`Dòng ${rowNumber}: Tên đăng nhập '${username}' đã tồn tại trong hệ thống.`);
          } else {
            errors.push(`Dòng ${rowNumber}: Email '${email}' đã được sử dụng trong hệ thống.`);
          }
          continue;
        }

        const password_hash = await bcrypt.hash(password, 10);
        await User.create({
          full_name,
          username,
          email,
          password_hash,
          role,
          status: 'Approved',
          student_staff_id,
          class_name: role === 'Student' ? class_name : (class_name || null),
          department,
          email_verified: true,
        });
        successful++;
      } catch (err) {
        failed++;
        errors.push(`Dòng ${rowNumber}: ${err.errors?.[0]?.message || 'Lỗi lưu trữ cơ sở dữ liệu'}`);
      }
    }

    res.json({
      message: `Nhập người dùng hoàn tất. Thành công: ${successful}, Thất bại: ${failed}.`,
      successful,
      failed,
      errors: errors.slice(0, 50),
    });
  } catch (err) {
    logger.error('importUsers error:', err);
    res.status(500).json({ message: 'Không thể xử lý file Excel.' });
  }
};

// Approves (`Approved`), rejects (`Rejected`), or locks (`Locked`) user accounts.
exports.updateUserStatus = async (req, res) => {
  try {
    const { id } = req.params;
    const { status, reject_reason } = req.body;

    if (!['Approved', 'Rejected', 'Locked'].includes(status)) {
      return res.status(400).json({ message: 'Trạng thái không hợp lệ. Phải là Approved, Rejected hoặc Locked.' });
    }

    const user = await User.findByPk(id, { attributes: ['id', 'full_name', 'email', 'status', 'role'] });
    if (!user) return res.status(404).json({ message: 'User not found.' });
    if (user.role === 'Admin') return res.status(403).json({ message: 'Cannot change admin account status this way.' });

    const isReasonNeeded = status === 'Rejected' || status === 'Locked';
    await user.update({
      status,
      reject_reason: isReasonNeeded ? (reject_reason || null) : null
    });

    let notifTitle = '';
    let notifMessage = '';

    if (status === 'Approved') {
      notifTitle = 'Tài khoản đã được phê duyệt';
      notifMessage = 'Tài khoản của bạn đã được Quản trị viên phê duyệt thành công. Bạn hiện có thể đăng nhập và trải nghiệm toàn bộ tính năng của EcoSurvey.';
    } else if (status === 'Locked') {
      notifTitle = 'Tài khoản đã bị khóa';
      notifMessage = `Tài khoản của bạn đã bị Quản trị viên khóa. ${reject_reason ? 'Lý do: ' + reject_reason : 'Vui lòng liên hệ Quản trị viên để biết thêm chi tiết.'}`;
    } else {
      notifTitle = 'Đơn đăng ký bị từ chối';
      notifMessage = `Yêu cầu đăng ký tài khoản của bạn đã bị từ chối. ${reject_reason ? 'Lý do: ' + reject_reason : 'Vui lòng liên hệ Quản trị viên để biết thêm chi tiết.'}`;
    }

    await Notification.create({
      user_id:        user.id,
      title:          notifTitle,
      message:        notifMessage,
      reference_type: 'user',
      reference_id:   user.id,
    });

    emailService.sendStatusUpdateEmail(user.email, user.full_name, status, reject_reason).catch(logger.error);

    res.json({ message: `Cập nhật trạng thái tài khoản thành ${status} thành công.`, user });
  } catch (err) {
    logger.error('updateUserStatus error:', err);
    res.status(500).json({ message: 'Failed to update user status.' });
  }
};

// Soft-deletes user account by setting status to Deactivated.
exports.deleteUser = async (req, res) => {
  try {
    const { id } = req.params;
    if (parseInt(id) === req.user.id) {
      return res.status(400).json({ message: 'You cannot deactivate your own account.' });
    }
    const user = await User.findByPk(id, { attributes: ['id', 'full_name', 'email', 'status', 'role'] });
    if (!user) return res.status(404).json({ message: 'User not found.' });
    if (user.role === 'Admin') return res.status(403).json({ message: 'Cannot deactivate admin accounts.' });
    if (user.status === 'Deactivated') {
      return res.status(400).json({ message: 'User is already deactivated.' });
    }

    await user.update({ status: 'Deactivated' });
    res.json({ message: 'User has been deactivated. Their historical data is preserved.' });
  } catch (err) {
    logger.error('deleteUser error:', err);
    res.status(500).json({ message: 'Failed to deactivate user.' });
  }
};

// Aggregates system-wide overview statistics.
exports.getStats = async (req, res) => {
  try {
    const [usersByRole, usersByStatus, surveysByStatus, recentParticipations] = await Promise.all([
      User.findAll({
        attributes: ['role', [User.sequelize.fn('COUNT', User.sequelize.col('id')), 'count']],
        group: ['role'],
        raw: true,
      }),
      User.findAll({
        attributes: ['status', [User.sequelize.fn('COUNT', User.sequelize.col('id')), 'count']],
        group: ['status'],
        raw: true,
      }),
      Survey.findAll({
        attributes: ['status', [Survey.sequelize.fn('COUNT', Survey.sequelize.col('id')), 'count']],
        group: ['status'],
        raw: true,
      }),
      SurveyResponse.count({
        where: {
          submitted_at: { [Op.gte]: new Date(Date.now() - 7 * 86400 * 1000) },
        },
      }),
    ]);

    const pendingParticipations = await Participation.count({ where: { status: 'Pending' } });
    const totalUsers = await User.count();

    res.json({ usersByRole, usersByStatus, surveysByStatus, recentParticipations, pendingParticipations, totalUsers });
  } catch (err) {
    logger.error('getStats error:', err);
    res.status(500).json({ message: 'Failed to fetch stats.' });
  }
};

// Retrieves list of pending extracurricular proof reports (`status = Pending`).
exports.getPendingParticipations = async (req, res) => {
  try {
    const { page, limit, offset } = getPagination(req.query);

    const { count, rows } = await Participation.findAndCountAll({
      where: { status: 'Pending' },
      include: [{ model: User, as: 'user', attributes: ['id', 'full_name', 'username', 'role'] }],
      order: [['created_at', 'DESC']],
      limit,
      offset,
    });

    res.json({ participations: rows, total: count, page, totalPages: Math.ceil(count / limit) });
  } catch (err) {
    logger.error('getPendingParticipations error:', err);
    res.status(500).json({ message: 'Failed to fetch pending participations.' });
  }
};

// Retrieves list of active FAQs for admin management.
exports.getFAQs = async (_req, res) => {
  try {
    const faqs = await FAQ.findAll({
      where: { is_active: true },
      order: [['created_at', 'DESC']],
    });
    res.json({ faqs });
  } catch (err) {
    res.status(500).json({ message: 'Failed to fetch FAQs.' });
  }
};

// Adds a new FAQ entry.
exports.createFAQ = async (req, res) => {
  try {
    const { question, answer, category } = req.body;
    if (!question || !answer) return res.status(400).json({ message: 'Question and answer are required.' });
    const faq = await FAQ.create({ question, answer, category });
    res.status(201).json({ message: 'FAQ created.', faq });
  } catch (err) {
    res.status(500).json({ message: 'Failed to create FAQ.' });
  }
};

// Updates FAQ content.
exports.updateFAQ = async (req, res) => {
  try {
    const faq = await FAQ.findByPk(req.params.id);
    if (!faq) return res.status(404).json({ message: 'FAQ not found.' });

    const { question, answer, category, is_active } = req.body;
    await faq.update({ question, answer, category, is_active });
    res.json({ message: 'FAQ updated.', faq });
  } catch (err) {
    res.status(500).json({ message: 'Failed to update FAQ.' });
  }
};

// Deletes an FAQ entry.
exports.deleteFAQ = async (req, res) => {
  try {
    const faq = await FAQ.findByPk(req.params.id);
    if (!faq) return res.status(404).json({ message: 'FAQ not found.' });
    await faq.destroy();
    res.json({ message: 'FAQ deleted.' });
  } catch (err) {
    res.status(500).json({ message: 'Failed to delete FAQ.' });
  }
};

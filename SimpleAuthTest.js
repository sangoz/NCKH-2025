/**
 * SIMPLE AUTH TEST - Chạy function này để trigger authorization popup
 * Function cực kỳ đơn giản chỉ để force Google hiện popup
 */
function simpleAuthTest() {
  // Đơn giản nhất có thể - chỉ 1 dòng gọi MailApp
  MailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Test Authorization',
    'If you receive this email, authorization is working!'
  );

  Logger.log('Email sent successfully!');
}

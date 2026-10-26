package vn.mercury.cdr;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.GraphicsEnvironment;
import java.awt.MouseInfo;
import java.awt.Point;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.InputMap;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;
import javax.swing.WindowConstants;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AutoClicker {

  // ====== CẤU HÌNH ======
  static final String EXCEL_FILE = "data.xlsx"; // file Excel (phải đóng trong Excel khi chạy)
  static final String STOP_FILE = "stop.txt";   // tạo file này để dừng
  static final int DELAY_MS = 1000;
  static final int START_DELAY_SEC = 5;         // đếm ngược trước khi bắt đầu chạy
  static final int CAPTURE_SEC = 5;             // thời gian chờ khi lấy mỗi toạ độ

  // ====== NHẬN DIỆN CARD TÊN (vị trí 3) ======
  static final boolean AUTO_DETECT = true;      // true: tự kiểm tra card có xuất hiện ở vị trí 3 không
  static final boolean ASK_USER = false;        // true: sau khi click vị trí 3 hỏi C/S (false = tự chạy hẳn)
  static final int HOVER_WAIT_MS = 500;         // đợi bao lâu sau khi hover rồi mới đọc màu
  static final int WHITE_MIN = 250;             // R, G, B đều >= số này thì coi là màu trắng (trắng = 255, xám card hover ~ 241)

  // Mac dán bằng Cmd, Windows/Linux dán bằng Ctrl
  static final boolean IS_MAC = System.getProperty("os.name", "").toLowerCase().contains("mac");

  static Robot robot;
  static final DataFormatter fmt = new DataFormatter();

  enum Choice { CONTINUE, SKIP, STOP }

  public static void main(String[] args) throws Exception {
    robot = new Robot();

    // Chọn chức năng: "dattoado" hoặc "chay"
    String mode = args.length > 0 ? args[0] : "chay";

//    if (mode.equalsIgnoreCase("dattoado")) {
//      datToaDo();
//    } else {
//      chayTuDong();
//    }
    if (true) {
      chayTuDong();
    } else {
      chayTuDong();
    }
    System.exit(0); // đảm bảo JVM thoát (cửa sổ Swing có thể giữ tiến trình sống)
  }

  // ================== DATTOADO ==================
  // Lưu 3 vị trí chuột: vị trí 1 -> B3,C3 ; vị trí 2 -> B4,C4 ; vị trí 3 -> B5,C5
  static void datToaDo() throws Exception {
    try (FileInputStream in = new FileInputStream(EXCEL_FILE); Workbook wb = new XSSFWorkbook(in)) {

      Sheet sh = wb.getSheetAt(0);

      // Xóa B3:C5
      for (int r = 2; r <= 4; r++) {
        for (int c = 1; c <= 2; c++) {
          getCell(sh, r, c).setBlank();
        }
      }

      for (int k = 0; k < 3; k++) {
        Point p = capture(k + 1);
        getCell(sh, 2 + k, 1).setCellValue(p.x); // cột B
        getCell(sh, 2 + k, 2).setCellValue(p.y); // cột C
      }

      saveWorkbook(wb);
      System.out.println("Da luu toa do vao B3:C5.");
    }
  }

  static Point capture(int n) throws InterruptedException {
    System.out.println("Di chuot den vi tri " + n + " va de yen (" + CAPTURE_SEC + " giay)...");
    for (int s = CAPTURE_SEC; s > 0; s--) {
      System.out.println("  " + s + "...");
      Thread.sleep(1000);
    }
    Point p = MouseInfo.getPointerInfo().getLocation();
    System.out.println("  -> vi tri " + n + " = (" + p.x + ", " + p.y + ")");
    return p;
  }

  // ================== CHAYTUDONG ==================
  static void chayTuDong() throws Exception {
    try (FileInputStream in = new FileInputStream(EXCEL_FILE); Workbook wb = new XSSFWorkbook(in)) {

      Sheet sh = wb.getSheetAt(0);

      int x1 = readInt(sh, 2, 1); // B3
      int y1 = readInt(sh, 2, 2); // C3
      int x2 = readInt(sh, 3, 1); // B4
      int y2 = readInt(sh, 3, 2); // C4
      int x3 = readInt(sh, 4, 1); // B5
      int y3 = readInt(sh, 4, 2); // C5

      if ((x1 == 0 && y1 == 0) || (x2 == 0 && y2 == 0) || (x3 == 0 && y3 == 0)) {
        System.out.println("Chua co du 3 toa do trong B3:C5. Hay chay 'dattoado' truoc.");
        return;
      }

      // Nội dung G1:G20 (ghép bằng xuống dòng)
      StringBuilder sb = new StringBuilder();
      for (int r = 0; r < 20; r++) {
        sb.append(fmt.formatCellValue(getCell(sh, r, 6)));
        if (r < 19)
          sb.append("\n");
      }
      String blockG = sb.toString();
      System.out.println("Noi dung G1:G20: " + blockG.replace("\n", "").length() + " ky tu.");
      if (blockG.trim().isEmpty()) {
        System.out.println("G1:G20 dang trong -> khong co noi dung de dan. Dung.");
        return;
      }

      // ===== ĐẾM NGƯỢC =====
      System.out.println("Bat dau sau " + START_DELAY_SEC + " giay. Chuyen sang ung dung dich...");
      for (int s = START_DELAY_SEC; s > 0; s--) {
        System.out.println(s + "...");
        Thread.sleep(1000);
      }
      System.out.println("Bat dau!");

      int lastRow = sh.getLastRowNum(); // chỉ số 0-based của dòng cuối

      // i = 1 tương ứng dòng 2 trong Excel
      for (int i = 1; i <= lastRow; i++) {

        if (new File(STOP_FILE).exists()) {
          System.out.println("Phat hien stop.txt -> dung.");
          return;
        }

        String phone = fmt.formatCellValue(getCell(sh, i, 4)); // cột E
        String done = fmt.formatCellValue(getCell(sh, i, 5)); // cột F

        if (phone.isEmpty() || !done.isEmpty())
          continue;

        // 1) Copy số điện thoại vào clipboard (chưa ghi gì vào cột F)
        setClipboard(phone);
        sleep(DELAY_MS);

        try {
          // 2) Di chuyển chuột tới vị trí 1 và double click
          robot.mouseMove(x1, y1);
          sleep(DELAY_MS);
          click();
          click();

          sleep(DELAY_MS);
          pressKey(KeyEvent.VK_DELETE);
          sleep(DELAY_MS);

          ctrlV(); // dán số điện thoại
          sleep(DELAY_MS);

          // Hover vị trí 3, đợi 0.5s: màu trắng -> không có card -> bỏ qua; khác trắng -> gửi tin
          if (AUTO_DETECT && !hoverKhacTrang(x3, y3, "vi tri 3 (card)")) {
            markRow(wb, sh, i, "failed");
            System.out.println("failed (khong thay card) so " + phone + " (dong " + (i + 1) + ").");
            continue;
          }

          // 3) Click 2 lần (double click) vào vị trí 3 (thay cho Enter trước đây)
          robot.mouseMove(x3, y3);
          sleep(DELAY_MS);
          click();
          sleep(DELAY_MS);

          // Sau khi bấm card: hover vị trí 2 (ô soạn tin). Khác trắng -> không gửi tin được -> bỏ qua số này
          if (AUTO_DETECT && hoverKhacTrang(x2, y2, "vi tri 2 (o nhan tin)")) {
            markRow(wb, sh, i, "failed");
            System.out.println("failed (o nhan tin khong trang) so " + phone + " (dong " + (i + 1) + ").");

            robot.mouseMove(x3, y3); // giống nhánh bấm S: click vị trí 3 lại trước khi sang số kế tiếp
            sleep(DELAY_MS);
            click();
            continue;
          }

          // ===== TẠM DỪNG (sau khi click vị trí 3): C = tiếp tục, S = bỏ qua số này =====
          Choice choice = ASK_USER ? askUser(phone) : Choice.CONTINUE;
          if (choice == Choice.STOP) {
            // Chưa xử lý xong, cột F để trống nên lần sau sẽ làm lại số này
            System.out.println("Nguoi dung chon dung. Dong " + (i + 1) + " chua danh dau.");
            return;
          }
          if (choice == Choice.SKIP) {
            markRow(wb, sh, i, "failed");
            System.out.println("failed (bo qua) so " + phone + " (dong " + (i + 1) + ").");

            robot.mouseMove(x3, y3);
            sleep(DELAY_MS);
            click();
            continue;
          }

          // 4) Copy G1:G20
          setClipboard(blockG);
          sleep(DELAY_MS);

          // 5) Di chuyển tới vị trí 2, click, dán
          robot.mouseMove(x2, y2);
          sleep(DELAY_MS);
          click(); // lần 1: đưa ứng dụng đích trở lại phía trước (sau khi cửa sổ hỏi đóng)
          sleep(DELAY_MS);
          click(); // lần 2: đặt con trỏ vào ô nhập để chắc chắn có focus
          sleep(DELAY_MS);
          ctrlV();
          sleep(DELAY_MS);
          pressKey(KeyEvent.VK_ENTER);
          pressKey(KeyEvent.VK_ENTER);

          // 6) Chạy đủ các bước -> ghi "ok"
          markRow(wb, sh, i, "ok");
          System.out.println("ok: " + phone + " (dong " + (i + 1) + ")");
        } catch (Exception e) {
          // Lỗi giữa chừng -> chưa xử lý đầy đủ -> ghi "failed" rồi dừng
          try {
            markRow(wb, sh, i, "failed");
          } catch (Exception ignore) {
            // bỏ qua lỗi ghi file, vẫn ném lỗi gốc
          }
          throw e;
        }
      }
      System.out.println("Hoan tat.");
    }
  }

  // ================== HỎI NGƯỜI DÙNG (C / S) ==================
  // Hiện một cửa sổ nhỏ ở góc phải dưới màn hình, luôn nằm trên cùng.
  // Phím C = tiếp tục, S = bỏ qua số này, Esc hoặc đóng cửa sổ = dừng hẳn.
  // Vì cửa sổ này nhận phím nên ký tự "c"/"s" KHÔNG bị gõ vào ứng dụng đích.
  static Choice askUser(String phone) throws Exception {
    final BlockingQueue<Choice> queue = new ArrayBlockingQueue<>(1);
    final JFrame[] holder = new JFrame[1];

    SwingUtilities.invokeAndWait(() -> {
      JFrame f = new JFrame("AutoClicker");
      f.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
      f.setAlwaysOnTop(true);
      f.addWindowListener(new WindowAdapter() {
        @Override
        public void windowClosing(WindowEvent e) {
          queue.offer(Choice.STOP);
        }
      });

      JLabel label = new JLabel("<html><div style='text-align:center'>Da dan so va click vi tri 3: <b>" + phone
          + "</b><br>Bam <b>C</b> = tiep tuc &nbsp;|&nbsp; <b>S</b> = bo qua so nay"
          + "<br>(Esc = dung han)</div></html>", JLabel.CENTER);
      label.setBorder(BorderFactory.createEmptyBorder(12, 16, 4, 16));

      JButton bContinue = new JButton("Tiep tuc (C)");
      JButton bSkip = new JButton("Bo qua (S)");
      JButton bStop = new JButton("Dung (Esc)");
      // Không cho nút giữ focus để phím C/S luôn đi vào cửa sổ
      bContinue.setFocusable(false);
      bSkip.setFocusable(false);
      bStop.setFocusable(false);
      bContinue.addActionListener(e -> queue.offer(Choice.CONTINUE));
      bSkip.addActionListener(e -> queue.offer(Choice.SKIP));
      bStop.addActionListener(e -> queue.offer(Choice.STOP));

      JPanel buttons = new JPanel(new FlowLayout(FlowLayout.CENTER, 8, 8));
      buttons.add(bContinue);
      buttons.add(bSkip);
      buttons.add(bStop);

      f.setLayout(new BorderLayout());
      f.add(label, BorderLayout.CENTER);
      f.add(buttons, BorderLayout.SOUTH);

      // Phím tắt hoạt động khi cửa sổ đang giữ focus
      InputMap im = f.getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW);
      im.put(KeyStroke.getKeyStroke(KeyEvent.VK_C, 0), "continue");
      im.put(KeyStroke.getKeyStroke(KeyEvent.VK_S, 0), "skip");
      im.put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), "stop");
      f.getRootPane().getActionMap().put("continue", new AbstractAction() {
        public void actionPerformed(ActionEvent e) { queue.offer(Choice.CONTINUE); }
      });
      f.getRootPane().getActionMap().put("skip", new AbstractAction() {
        public void actionPerformed(ActionEvent e) { queue.offer(Choice.SKIP); }
      });
      f.getRootPane().getActionMap().put("stop", new AbstractAction() {
        public void actionPerformed(ActionEvent e) { queue.offer(Choice.STOP); }
      });

      f.pack();
      // Đặt ở góc phải dưới để không che ứng dụng đích
      Rectangle b = GraphicsEnvironment.getLocalGraphicsEnvironment().getMaximumWindowBounds();
      f.setLocation(b.x + b.width - f.getWidth() - 20, b.y + b.height - f.getHeight() - 20);
      f.setVisible(true);
      f.toFront();
      f.requestFocus();
      Toolkit.getDefaultToolkit().beep();
      holder[0] = f;
    });

    Choice choice = queue.take();

    // Đóng cửa sổ trước khi robot click tiếp để không che vị trí cần click
    SwingUtilities.invokeAndWait(() -> holder[0].dispose());
    Thread.sleep(400);
    return choice;
  }

  // ================== NHẬN DIỆN CARD (theo màu điểm ảnh) ==================
  // Màu có được coi là trắng không (cả R, G, B >= WHITE_MIN)
  static boolean isWhite(Color c) {
    return c.getRed() >= WHITE_MIN && c.getGreen() >= WHITE_MIN && c.getBlue() >= WHITE_MIN;
  }

  // Hover tại (x, y), đợi HOVER_WAIT_MS rồi đọc màu 1 lần. Khác trắng -> true.
  static boolean hoverKhacTrang(int x, int y, String ten) {
    // Rung chuột 1 pixel để ứng dụng chắc chắn cập nhật trạng thái hover
    robot.mouseMove(x + 1, y);
    robot.delay(100);
    robot.mouseMove(x, y);
    robot.delay(HOVER_WAIT_MS);

    Color c = robot.getPixelColor(x, y);
    boolean khacTrang = !isWhite(c);
    System.out.printf("  mau tai %s: RGB(%d, %d, %d) -> %s%n", ten, c.getRed(), c.getGreen(), c.getBlue(),
        khacTrang ? "khac trang" : "trang");
    return khacTrang;
  }

  // ================== HÀM TIỆN ÍCH ==================
  static void click() {
    robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
    robot.delay(80);
    robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    robot.delay(120);
  }

  static void pressKey(int keyCode) {
    robot.keyPress(keyCode);
    robot.delay(30);
    robot.keyRelease(keyCode);
    robot.delay(50);
  }

  static void ctrlV() {
    int mod = IS_MAC ? KeyEvent.VK_META : KeyEvent.VK_CONTROL;
    robot.keyPress(mod);
    robot.keyPress(KeyEvent.VK_V);
    robot.delay(30);
    robot.keyRelease(KeyEvent.VK_V);
    robot.keyRelease(mod);
    robot.delay(50);
  }

  static void setClipboard(String text) {
    Toolkit.getDefaultToolkit().getSystemClipboard().setContents(new StringSelection(text), null);
  }

  static void sleep(int ms) throws InterruptedException {
    Thread.sleep(ms);
  }

  // Ghi trạng thái vào cột F của dòng i rồi lưu file
  static void markRow(Workbook wb, Sheet sh, int i, String status) throws Exception {
    getCell(sh, i, 5).setCellValue(status);
    saveWorkbook(wb);
  }

  static int readInt(Sheet sh, int row, int col) {
    Cell c = getCell(sh, row, col);
    if (c.getCellType() == CellType.NUMERIC)
      return (int) c.getNumericCellValue();
    return 0;
  }

  static Cell getCell(Sheet sh, int row, int col) {
    Row r = sh.getRow(row);
    if (r == null)
      r = sh.createRow(row);
    Cell c = r.getCell(col);
    if (c == null)
      c = r.createCell(col);
    return c;
  }

  static void saveWorkbook(Workbook wb) throws Exception {
    try (FileOutputStream out = new FileOutputStream(EXCEL_FILE)) {
      wb.write(out);
    }
  }
}
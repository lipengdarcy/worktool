package cn.zq.gui;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;
import cn.zq.tool.ExcelTool;

public class MainFrame extends JFrame {

	private static final long serialVersionUID = 1L;
	private String filePath;

	JLabel fileName;

	public MainFrame() {
		this.setTitle("周颀的工具");
		this.setSize(500, 300);
		JPanel panel = new JPanel();
		this.add(panel);
		placeComponents(panel);
		this.setVisible(true);
	}

	private void placeComponents(JPanel panel) {

		panel.setLayout(new GridLayout(3, 1));

		// 1.首先定义一个Button
		JButton button_file = new JButton("选择文件");
		panel.add(button_file);
		// 为该Button添加时间监听器，在监听器中加入文件选择器：
		button_file.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileFilter(new FileNameExtensionFilter("Excel文件(*.xls)", "xls"));
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择");
				File file = jfc.getSelectedFile();
				if (file != null) {
					filePath = file.getAbsolutePath();
					fileName.setText("所选文件: " + filePath);
					ExcelTool.processData(filePath);
					JOptionPane.showMessageDialog(null, "excel文件处理完毕，文件路径：" + filePath, "提醒",
							JOptionPane.YES_NO_OPTION);

				}
			}
		});

		// 2.加一个label，显示选择的文件全路径
		fileName = new JLabel();
		panel.add(fileName);

	}

}

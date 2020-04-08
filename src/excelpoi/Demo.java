package excelpoi;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JLabel;
import java.awt.Color;
import java.awt.Font;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.awt.event.ActionEvent;
import excelpoi.SimpleExcelReaderExample;
import javax.swing.LayoutStyle.ComponentPlacement;

public class Demo extends JFrame {

	private JPanel contentPane;
	private JTextField t2;
	private JTextField t1;
	private JTextField t3;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Demo frame = new Demo();
					frame.setVisible(true);
					
					
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Demo() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		
		JLabel lblNewLabel = new JLabel("Automation Datasheet Validation");
		lblNewLabel.setFont(new Font("Calibri", Font.BOLD, 14));
		lblNewLabel.setForeground(Color.BLACK);
		
		JLabel lblDatasheetPath = new JLabel("Datasheet Path");
		lblDatasheetPath.setFont(new Font("Tahoma", Font.PLAIN, 11));
		
		JLabel lblTestcaseName = new JLabel("Control Sheet");
		lblTestcaseName.setFont(new Font("Tahoma", Font.PLAIN, 11));
		
		t2 = new JTextField();
		t2.setColumns(10);
		
		t1 = new JTextField();
		t1.setColumns(10);
		
		JButton btnValidate = new JButton("Validate");
		btnValidate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				//SimpleExcelReaderExample bt = new SimpleExcelReaderExample();
				Excelreader bt = new Excelreader();
				
				
						bt.inputauto = t1.getText();
						System.out.println(bt.inputauto);
				        bt.inputcontrol = t2.getText();
				        System.out.println(bt.inputcontrol);
				        bt.inputtcNm = t2.getText();
				        System.out.println(bt.inputcontrol);
				
				        try {
							bt.excelexecute();
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
				        
				System.out.println("**************************************");
				
				System.out.println("complete");
			}
		});
		btnValidate.setBackground(new Color(175, 238, 238));
		btnValidate.setForeground(Color.BLACK);
		btnValidate.setFont(new Font("Tahoma", Font.BOLD, 11));
		
		t3 = new JTextField();
		t3.setColumns(10);
		
		JLabel label = new JLabel("Testcase Name");
		label.setFont(new Font("Tahoma", Font.PLAIN, 11));
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.TRAILING)
				.addGroup(Alignment.LEADING, gl_contentPane.createSequentialGroup()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addContainerGap()
							.addComponent(btnValidate))
						.addGroup(Alignment.LEADING, gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addComponent(lblTestcaseName)
								.addComponent(lblDatasheetPath, GroupLayout.PREFERRED_SIZE, 85, GroupLayout.PREFERRED_SIZE)
								.addComponent(label, GroupLayout.PREFERRED_SIZE, 83, GroupLayout.PREFERRED_SIZE))
							.addGap(23)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(Alignment.TRAILING, gl_contentPane.createSequentialGroup()
									.addComponent(t3, GroupLayout.DEFAULT_SIZE, 153, Short.MAX_VALUE)
									.addGap(141))
								.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING, false)
									.addComponent(lblNewLabel, GroupLayout.PREFERRED_SIZE, 211, GroupLayout.PREFERRED_SIZE)
									.addComponent(t1, GroupLayout.DEFAULT_SIZE, 294, Short.MAX_VALUE)
									.addComponent(t2)))))
					.addContainerGap(22, Short.MAX_VALUE))
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addComponent(lblNewLabel)
					.addGap(58)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblDatasheetPath)
						.addComponent(t1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(40)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblTestcaseName)
						.addComponent(t2, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(label)
						.addComponent(t3, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addComponent(btnValidate)
					.addContainerGap(14, Short.MAX_VALUE))
		);
		contentPane.setLayout(gl_contentPane);
	}
}

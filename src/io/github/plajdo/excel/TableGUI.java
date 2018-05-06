package io.github.plajdo.excel;

import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import net.miginfocom.swing.MigLayout;

public class TableGUI {

	private JFrame frmExcelStuff;
	private JTextField textField;
	private JButton btnChoosepath;
	private JLabel lblOutputlabel;
	private JTextField textField_1;
	private JButton btnNewButton;
	private JProgressBar progressBar;
	private JButton btnSpracova;
	
	private static TableGUI instance;
	private JLabel lblObrzok;
	private JTextField textField_2;
	private JButton btnSrlecc;
	private JLabel lblProtokolOKontrole;
	private JCheckBox chckbxucnonie;
	private JLabel lblHshzm;
	private JCheckBox chckbxAnonie;
	
	public static TableGUI getInstance(){
		return instance;
	}
	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					instance = new TableGUI();
					instance.frmExcelStuff.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
				
			}
			
		});
		
	}

	/**
	 * Create the application.
	 */
	private TableGUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmExcelStuff = new JFrame();
		frmExcelStuff.setTitle("Excel stuff");
		frmExcelStuff.setBounds(100, 100, 550, 300);
		frmExcelStuff.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmExcelStuff.getContentPane().setLayout(new MigLayout("", "[][grow][]", "[][][][][][grow][][grow][]"));
		
		JLabel lblKmexls = new JLabel("Kme\u0148.xls:");
		frmExcelStuff.getContentPane().add(lblKmexls, "cell 0 0,alignx trailing");
		
		textField = new JTextField();
		frmExcelStuff.getContentPane().add(textField, "flowx,cell 1 0,growx");
		textField.setColumns(10);
		
		btnChoosepath = new JButton("Vybra\u0165 s\u00FAbor");
		btnChoosepath.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				chooser.setAcceptAllFileFilterUsed(false);
				
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel 97-2003 Workbook", "xls");
				chooser.setFileFilter(filter);
				
				int returnVal = chooser.showOpenDialog(frmExcelStuff);
				
				if(returnVal == JFileChooser.APPROVE_OPTION){
					textField.setText(chooser.getSelectedFile().getPath());
				}
				
			}
			
		});
		frmExcelStuff.getContentPane().add(btnChoosepath, "cell 2 0,growx");
		
		lblObrzok = new JLabel("Obr\u00E1zok:");
		frmExcelStuff.getContentPane().add(lblObrzok, "cell 0 1,alignx trailing");
		
		textField_2 = new JTextField();
		frmExcelStuff.getContentPane().add(textField_2, "cell 1 1,growx");
		textField_2.setColumns(10);
		
		btnSrlecc = new JButton("Vybra\u0165 s\u00FAbor");
		btnSrlecc.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				chooser.setAcceptAllFileFilterUsed(false);
				
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Portable Network Graphics", "png");
				chooser.setFileFilter(filter);
				
				int returnVal = chooser.showOpenDialog(frmExcelStuff);
				
				if(returnVal == JFileChooser.APPROVE_OPTION){
					textField_2.setText(chooser.getSelectedFile().getPath());
				}
				
			}
			
		});
		frmExcelStuff.getContentPane().add(btnSrlecc, "cell 2 1,growx");
		
		lblOutputlabel = new JLabel("V\u00FDstupn\u00FD prie\u010Dinok:");
		frmExcelStuff.getContentPane().add(lblOutputlabel, "cell 0 2,alignx trailing");
		
		textField_1 = new JTextField();
		frmExcelStuff.getContentPane().add(textField_1, "flowx,cell 1 2,growx");
		textField_1.setColumns(10);
		
		btnNewButton = new JButton("Vybra\u0165 prie\u010Dinok");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);
				
				if(chooser.showOpenDialog(frmExcelStuff) == JFileChooser.APPROVE_OPTION){
					textField_1.setText(chooser.getSelectedFile().getPath() + File.separator);
				}
				
			}
			
		});
		frmExcelStuff.getContentPane().add(btnNewButton, "cell 2 2,growx");
		
		lblProtokolOKontrole = new JLabel("Protokol o kontrole:");
		frmExcelStuff.getContentPane().add(lblProtokolOKontrole, "cell 0 3");
		
		chckbxucnonie = new JCheckBox("\u00C1no/Nie");
		frmExcelStuff.getContentPane().add(chckbxucnonie, "flowx,cell 1 3");
		
		lblHshzm = new JLabel("Triedenie HS (pre CFH):");
		frmExcelStuff.getContentPane().add(lblHshzm, "cell 0 4");
		
		chckbxAnonie = new JCheckBox("\u00C1no/Nie");
		frmExcelStuff.getContentPane().add(chckbxAnonie, "cell 1 4");
		
		progressBar = new JProgressBar();
		progressBar.setStringPainted(true);
		frmExcelStuff.getContentPane().add(progressBar, "cell 0 6 3 1,growx");
		
		btnSpracova = new JButton("Spracova\u0165");
		btnSpracova.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Runnable r = () -> {
					try{
						if(!chckbxAnonie.isSelected()){
							FilterExcelTable.createHZM(new File(textField.getText()), new File(textField_2.getText()), textField_1.getText(), chckbxucnonie.isSelected());	
							JOptionPane.showMessageDialog(frmExcelStuff, "Dokon\u010Den\u00E9", "Hotovo", JOptionPane.INFORMATION_MESSAGE);
						}else{
							FilterExcelTable.createHS(new File(textField.getText()), new File(textField_2.getText()), textField_1.getText(), chckbxucnonie.isSelected());	
							JOptionPane.showMessageDialog(frmExcelStuff, "Dokon\u010Den\u00E9", "Hotovo", JOptionPane.INFORMATION_MESSAGE);
						}
					}catch(Exception e1){
						e1.printStackTrace();
						JOptionPane.showMessageDialog(frmExcelStuff, "Chyba pri spracovan\u00ED tabu\u013Eky! Popis chyby:\n" + e1.toString(), "Chyba", JOptionPane.ERROR_MESSAGE);
					}finally{
						textField.setText("");
						textField_1.setText("");
						textField_2.setText("");
						progressBar.setIndeterminate(false);
						progressBar.setValue(0);
					}

				};
				Thread t = new Thread(r);
				t.start();

			}

		});
		frmExcelStuff.getContentPane().add(btnSpracova, "cell 0 8");
		
	}
	
	public void setProgress(int pg){
		assert pg < 100 && pg > -1;
		if(pg == -1){
			progressBar.setIndeterminate(true);
		}else{
			progressBar.setIndeterminate(false);
			progressBar.setValue(pg);			
		}
		
	}

}

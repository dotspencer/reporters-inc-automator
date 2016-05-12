package reportersInc;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextArea;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import reportersInc.View;

public class View implements ActionListener{
	
	JFrame frame;
	JPanel top;
	JTextArea result;
	
	String folderMessage = "Folder created";
	String blankDocMessage = "Blank .doc file created";
	String textWrittenMessage = "Text written to .doc file";
	
	public View(){
		setUpFrame();
		setUpPanel();
		frame.setVisible(true);
	}
	
	private void setUpFrame(){
		frame = new JFrame();
		frame.setTitle("ASCII Converter");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(new Dimension(300, 150));
		frame.setMinimumSize(frame.getSize());
		frame.setLayout(new BorderLayout());
		frame.setLocationRelativeTo(null);
		
		result = new JTextArea();
		//result.setFont(new Font("Courier", 0, 14));
		result.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
		frame.add(result, BorderLayout.CENTER);
	}
	
	private void setUpPanel(){
		top = new JPanel();
		
		JButton fileButton = new JButton("Select File");
		top.add(fileButton);
		fileButton.addActionListener(this);
		
		frame.add(top, BorderLayout.NORTH);
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		result.setText("");
		
		JFileChooser chooser = new JFileChooser();
		int result = chooser.showOpenDialog(null);
		
		if(result != JFileChooser.APPROVE_OPTION){
			return;
		}
		
		File ascii = chooser.getSelectedFile();
		System.out.println(ascii.toString());
		
		if(!createFolder(ascii)){
			JOptionPane.showMessageDialog(null, "Folder already exists.");
		} else {
			
			createDoc(ascii); // Creates .doc file if folder was created
			// writeDoc() method is called inside createDoc()
			copyText(ascii);
		}
	}
	
	private boolean createFolder(File file){
		File dir = new File(file.getParent() + "/" + justName(file));
		boolean folderMade = dir.mkdir();
		if(folderMade){
			addResult(folderMessage);
		}
		return folderMade;
	}
	
	private void createDoc(File ascii){
		InputStream blank = View.class.getResourceAsStream("/reportersInc/blank.doc");
		Path target = new File(ascii.getParent() + "/" + justName(ascii) + "/" + justName(ascii) + ".doc").toPath();
		try {
			Files.copy(blank, target, StandardCopyOption.REPLACE_EXISTING);
			addResult(blankDocMessage);
		} catch (IOException e) {
			e.printStackTrace();
		}
		writeDoc(ascii, new File(target.toString()));
		addResult(textWrittenMessage);
	}
	
	private void writeDoc(File ascii, File dotDoc){
		FileReader fr = null;
		
		try {
			fr = new FileReader(ascii);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		HWPFDocument doc = null;
		
		try {
			FileInputStream in = new FileInputStream(dotDoc);
			doc = new HWPFDocument(in);
			//in.close();
			
			Range range = doc.getRange();
			
			while(fr.ready()){
				range.insertAfter((char)fr.read() + "");
			}
			
			FileOutputStream out = new FileOutputStream(dotDoc);
			doc.createInformationProperties();
			doc.write(out);
			out.close();
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private void copyText(File file){
		Path target = new File(file.getParent() + "/" + justName(file) + "/" + justName(file) + ".txt").toPath();
		try {
			Files.copy(file.toPath(), target);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private String justName(File file){
		String name = file.getName();
		return name.substring(0, name.length() - 4);
	}
	
	private void addResult(String message){
		result.append(message + "\n");
	}
}

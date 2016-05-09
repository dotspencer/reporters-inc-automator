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
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

public class View implements ActionListener{
	
	JFrame frame;
	JPanel top;
	
	public View(){
		setUpFrame();
		setUpPanel();
		frame.setVisible(true);
	}
	
	private void setUpFrame(){
		frame = new JFrame();
		frame.setTitle("Text to Doc");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(new Dimension(300, 150));
		frame.setMinimumSize(frame.getSize());
		frame.setLayout(new BorderLayout());
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
		JFileChooser chooser = new JFileChooser();
		int result = chooser.showOpenDialog(null);
		
		if(result != JFileChooser.APPROVE_OPTION){
			return;
		}
		
		File file = chooser.getSelectedFile();
		System.out.println(file.toString());
		
		if(!createFolder(file)){
			JOptionPane.showMessageDialog(null, "Folder already exists.");
		} else {
			createDoc(file); // Creates .doc file if folder was created
			// writeDoc() method is called inside createDoc()
			copyText(file);
		}
	}
	
	private boolean createFolder(File file){
		return new File(file.getParent() + "/" + justName(file)).mkdir();
	}
	
	private void createDoc(File file){
		Path target = new File(file.getParent() + "/" + justName(file) + "/" + justName(file) + ".doc").toPath();
		try {
			Files.copy(new File("Blank.doc").toPath(), target, StandardCopyOption.REPLACE_EXISTING);
		} catch (IOException e) {
			e.printStackTrace();
		}
		writeDoc(file, new File(target.toString()));
	}
	
	private void writeDoc(File text, File file){
		FileReader fr = null;
		
		try {
			fr = new FileReader(text);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		HWPFDocument doc = null;
		
		try {
			FileInputStream in = new FileInputStream(file);
			doc = new HWPFDocument(in);
			//in.close();
			
			Range range = doc.getRange();
			range.replaceText("", false);
			
			while(fr.ready()){
				range.insertAfter((char)fr.read() + "");
			}
			
			Range after = doc.getRange();
			int numParagraphs = after.numParagraphs();
			
			for(int i = 0; i < numParagraphs; i++){
				Paragraph paragraph = after.getParagraph(i);
				
				int charRuns = paragraph.numCharacterRuns();
				for(int j = 0; j < charRuns; j++){
					int size = 9;
					CharacterRun run = paragraph.getCharacterRun(j);
					run.setFontSize(size*2); // In half sizes.
					run.setFtcAscii(4);
				}
			}
			
			FileOutputStream out = new FileOutputStream(file);
			doc.write(out);
			//out.close();
			
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
}

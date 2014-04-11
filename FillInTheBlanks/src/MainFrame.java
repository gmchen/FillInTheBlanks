import java.awt.AWTEvent;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.AWTEventListener;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Queue;
import java.util.Random;
import java.util.Scanner;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.microsoft.OfficeParser;
import org.apache.tika.parser.microsoft.ooxml.OOXMLParser;
import org.apache.tika.parser.odf.OpenDocumentParser;
import org.apache.tika.parser.pdf.PDFParser;
import org.apache.tika.parser.txt.TXTParser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.SAXException;

/**
 * @author Gregory M Chen
 */
public class MainFrame extends JFrame
{
	private List<String> phrases;
	private List<String> wordsToExclude;
	JButton fileDialogButton;
	JButton submitButton;
	JTextArea outputTextArea;
	JTextField inputTextField;
	JPanel bottomPanel;
	JPanel topPanel;
	JCheckBox slideshowCheckBox;
	Random random;
	/**
	 * Constructor for the MainFrame.
	 */
	public MainFrame() 
	{	
		phrases = new ArrayList<String>();
		wordsToExclude = new ArrayList<String>();
		random = new Random();
		getWordsToExlude();
		// Create a mouse click event listener
		long eventMask = AWTEvent.MOUSE_EVENT_MASK;
		Toolkit.getDefaultToolkit().addAWTEventListener( new AWTEventListener()
		{
		    public void eventDispatched(AWTEvent e)
		    {
		    	handleMouseEvent((MouseEvent) e);
		    }
		}, eventMask);
		
		this.setTitle("Questions From Slides");
		this.setSize(600, 400);
		this.setLocationRelativeTo(null);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setLayout(new BorderLayout(15,15));
		
		// Create file dialog button
		fileDialogButton = new JButton("Open File");
		fileDialogButton.setActionCommand("Open file dialog");
		fileDialogButton.addActionListener( new ActionListener()
		{
			public void actionPerformed(ActionEvent event) 
			{
				openDialog();
			}
		});
		
		// Create the slides/notes CheckBox
		slideshowCheckBox = new JCheckBox();
		slideshowCheckBox.setText("Only use text in notes for ppt/pptx files");
		slideshowCheckBox.setSelected(true);
		
		// Create the output area
		outputTextArea = new JTextArea("Output text area");
		outputTextArea.setVisible(true);
		outputTextArea.setBackground(new Color(98, 145, 255));
		
		// Create the input area
		inputTextField = new JTextField("Input text field");
		
		// Create the submit button
		submitButton = new JButton("Submit");
		submitButton.setActionCommand("Submit");
		submitButton.addActionListener( new ActionListener()
		{
			public void actionPerformed(ActionEvent event) 
			{
				submit();
			}
		});
		this.getRootPane().setDefaultButton(submitButton);
		
		// Create the top panel
		topPanel = new JPanel(new GridLayout());
		topPanel.add(fileDialogButton);
		topPanel.add(slideshowCheckBox);
		
		// Create a panel for the input area and submit button
		bottomPanel = new JPanel(new GridBagLayout());
		GridBagConstraints c1 = new GridBagConstraints();
		c1.fill = GridBagConstraints.HORIZONTAL;
		c1.gridwidth = 2;
		c1.gridx = 0;
		c1.gridy = 0;
		c1.weightx = 0.7;
		bottomPanel.add(inputTextField,c1);
		GridBagConstraints c2 = new GridBagConstraints();
		c2.fill = GridBagConstraints.HORIZONTAL;
		c2.gridwidth = 1;
		c2.gridx = 2;
		c2.gridy = 0;
		bottomPanel.add(submitButton,c2);
		
		// Add things to this JFrame
		this.add(topPanel, BorderLayout.NORTH);
		this.add(outputTextArea, BorderLayout.CENTER);
		this.add(bottomPanel, BorderLayout.SOUTH);
		
		// Start stuff
		this.setVisible(true);
		inputTextField.selectAll();
		inputTextField.requestFocus();
	}
	
	/**
	 * Open a file dialog and add the text.
	 */
	private void openDialog() {
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setMultiSelectionEnabled(true);
		fileChooser.setFileFilter(new FileNameExtensionFilter("Available file types", "ppt", "pptx", "doc", "docx", "odt", "odp", "txt", "pdf"));
		int returnVal = fileChooser.showOpenDialog(this);
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			File[] files = fileChooser.getSelectedFiles(); 
			for(File f:files) {
				String fileName = f.getAbsolutePath();
				String text = null;
				// Get contents if the file is doc, docx, odt, odp, or txt
				Parser parser = null;
				
				if(fileName.endsWith(".doc")) {
					parser = new OfficeParser();
				}
				else if(fileName.endsWith("docx")) {
					parser = new OOXMLParser();
				}
				else if(fileName.endsWith(".odt") || fileName.endsWith("odp")) {
					parser = new OpenDocumentParser();
				}
				else if(fileName.endsWith(".txt")) {
					parser = new TXTParser();
				}
				else if(fileName.endsWith(".pdf")) {
					parser = new PDFParser();
				}
				
				if(parser != null) {
					try {
						BodyContentHandler handler = new BodyContentHandler();
						parser.parse(new FileInputStream(f), handler, new Metadata(), new ParseContext());
						
						text = handler.toString();
					} catch (IOException | SAXException
							| TikaException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
				
				
				if(fileName.endsWith(".ppt"))
				{
					PowerPointExtractor powerpointExtractor = null;
					try {
						powerpointExtractor = new PowerPointExtractor(fileName);
					} catch (IOException e) {
						e.printStackTrace();
					}
					if(slideshowCheckBox.isSelected()) {
						text = powerpointExtractor.getNotes();
					}
					else {
						text = powerpointExtractor.getText();
						text += "\n" + powerpointExtractor.getNotes();
					}
				}
				if(fileName.endsWith(".pptx")) {
					XSLFPowerPointExtractor powerpointExtractor = null;
					try {
						powerpointExtractor = new XSLFPowerPointExtractor(new XMLSlideShow(new FileInputStream(fileName)));
					} catch (IOException e) {
						e.printStackTrace();
					}
					if(slideshowCheckBox.isSelected()) {
						text = powerpointExtractor.getText(false, true);
					}
					else {
						text = powerpointExtractor.getText(true, true);
					}
				}
				//if(fileName.endsWith(".doc") || fileName.endsWith(".docx")) {
					//ord6Extractor wordExtractor = new Word6Extractor(fileName);
				//}
				//System.out.println(text);
				String[] candidatePhrases = text.split("\n");
				for(String s:candidatePhrases) {
					if(s.length() < 2) {
						continue;
					}
					else if(s.length() < 30) {
						continue;
					}
					else {
						ArrayList<String> newPhrases = chopToSentences(s, 2, 300);
						phrases.addAll(newPhrases);
					}
				}
			}
		}
	}
	
	/**
	 * Attempt to chop a sentence into pieces less than maxLength, separated by periods.
	 */
	private ArrayList<String> chopToSentences(String largerString, int minLength, int maxLength) {
		Queue<String> q = new LinkedList<String>();
		ArrayList<String> finalSentenceList = new ArrayList<String>();
		q.add(largerString);
		while(!q.isEmpty()) {
			String str = q.poll();
			str = str.trim();
			str = str.replaceAll("\\•", "");
			str = str.replaceAll("^[\\- ]*", "");
			str = str.replaceAll(" +", " ");
			
			if(str.length() <= maxLength) {
				if(str.length() >= minLength) {
					finalSentenceList.add(str);
				}
			}
			else {
				// See if we can chop this into sub sentences
				List<Integer> periodIndices = new ArrayList<Integer>();
				int index = 0;
				while(index >= 0) {
					index = str.indexOf(".", index+1);
					if(index >= 0) {
						periodIndices.add(index);
					}
				}
				if(periodIndices.size() == 0) {
					finalSentenceList.add(str);
				}
				else {
					// Split in two and add pieces to the queue.
					int minDelta = Integer.MAX_VALUE;
					int bestIndex = -1;
					for(int i:periodIndices) {
						if( Math.abs(str.length() / 2 - i) < minDelta)  {
							minDelta = Math.abs(str.length()/2 - i);
							bestIndex = i;
						}
					}
					
					// If the period is not at the end
					if(bestIndex != str.length()-1) {
						q.add(str.substring(0, bestIndex+1));
						q.add(str.substring(bestIndex+1));
					}
					else {
						// If the period is at the end, then add this to the final list.
						finalSentenceList.add(str);
					}
				}
			}
		}
		
		return finalSentenceList;
	}

	
	/**
	 * Submit
	 */
	private void submit() {
		inputTextField.requestFocus();
		inputTextField.selectAll();
		
		// Get new phrase
		int index = random.nextInt(phrases.size());
		String phrase = phrases.get(index);
		String[] words = phrase.split(" ");
		ArrayList<String> removableWords = new ArrayList<String>();
		for(String word:words) {
			String compareWord = word.replaceAll("[^A-Za-z]", "").toUpperCase();
			int wordIndex = Collections.binarySearch(wordsToExclude, compareWord);
			if(wordIndex >= 0) {
				System.out.println(word + "\t" + wordsToExclude.get(wordIndex));
			}
			else {
				System.out.println(word);
			}
		}
	}
	
	/**
	 * Handle any mouse event by giving the input text area the focus
	 * @param e
	 */
	private void handleMouseEvent(MouseEvent e) {
		// No matter what, give the input text area the focus
		inputTextField.requestFocus();
	}
	
	private void getWordsToExlude() {
		Scanner scanner = null;
		try {
			scanner = new Scanner(new File("data/words_to_exclude.txt"));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		while(scanner.hasNext()) {
			wordsToExclude.add(scanner.next().toUpperCase());
		}
		Collections.sort(wordsToExclude, String.CASE_INSENSITIVE_ORDER);
	}
}

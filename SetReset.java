package forhonor.skillsetreset;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;

import java.awt.Font;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JTextArea;
import javax.swing.SwingConstants;
import javax.swing.JRadioButton;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.awt.event.ActionEvent;
import java.awt.Color;

@SuppressWarnings("serial")
public class SetReset extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					SetReset frame = new SetReset(null, null);
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 * 
	 * @param password
	 * @param username
	 */
	public SetReset(String username, String password) {
		setTitle("Skill SET/RESET Tool");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel lblSandbox = new JLabel("SANDBOX :");
		lblSandbox.setHorizontalAlignment(SwingConstants.TRAILING);
		lblSandbox.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblSandbox.setBounds(25, 54, 94, 14);
		contentPane.add(lblSandbox);

		JComboBox<String> comboBox = new JComboBox<String>();
		comboBox.setFont(new Font("Tahoma", Font.PLAIN, 13));
		comboBox.setBounds(129, 51, 261, 20);
		contentPane.add(comboBox);

		// Adding Sandoox to the Dropdown in the UI
		comboBox.addItem("HERO_PC_UAT_AA");
		comboBox.addItem("HERO_PC_UAT_T");
		comboBox.addItem("HERO_PC_UAT_F");
		comboBox.addItem("HERO_XBOXONE_UAT_O");
		comboBox.addItem("HERO_XBOXONE_UAT_D");
		comboBox.addItem("HERO_PS4_UAT_C");
		comboBox.addItem("HERO_PS4_UAT_M");

		JLabel lblImportSheet = new JLabel("Import Sheet :");
		lblImportSheet.setHorizontalAlignment(SwingConstants.TRAILING);
		lblImportSheet.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblImportSheet.setBounds(25, 112, 94, 14);
		contentPane.add(lblImportSheet);

		JLabel lbWelcome = new JLabel(username);
		lbWelcome.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lbWelcome.setBounds(156, 11, 268, 29);
		contentPane.add(lbWelcome);

		JTextArea textArea = new JTextArea();
		textArea.setToolTipText("Drop the Excel Sheet Here or Browse");
		textArea.setBounds(129, 108, 261, 52);
		contentPane.add(textArea);

		// Creating a Group for Radio Buttons SET and RESET
		ButtonGroup Radio = new ButtonGroup();

		JRadioButton rdbtnResetSkillRating = new JRadioButton("RESET Skill Rating", true);
		rdbtnResetSkillRating.setBounds(275, 183, 129, 23);
		Radio.add(rdbtnResetSkillRating);
		contentPane.add(rdbtnResetSkillRating);

		JRadioButton rdbtnSetSkillRating = new JRadioButton("SET Skill Rating");
		rdbtnSetSkillRating.setBounds(275, 209, 129, 23);
		Radio.add(rdbtnSetSkillRating);
		contentPane.add(rdbtnSetSkillRating);

		JButton btnBrowse = new JButton("Browse");
		btnBrowse.setForeground(new Color(0, 0, 139));
		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				// Browse excel File
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				fileChooser.setDialogTitle("Select Excel File");
				int rValue = fileChooser.showSaveDialog(null);
				if (rValue == JFileChooser.APPROVE_OPTION) {
					textArea.setText(fileChooser.getSelectedFile().getAbsolutePath());
				}
			}
		});
		btnBrowse.setFont(new Font("Tahoma", Font.PLAIN, 13));
		btnBrowse.setBounds(30, 137, 89, 23);
		contentPane.add(btnBrowse);

		JCheckBox chckbxResetAllBoards = new JCheckBox("RESET the database for all boards");
		chckbxResetAllBoards.setForeground(new Color(255, 0, 0));
		chckbxResetAllBoards.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (chckbxResetAllBoards.isSelected()) {
					lblImportSheet.setVisible(false);
					textArea.setVisible(false);
					btnBrowse.setVisible(false);
					rdbtnResetSkillRating.setVisible(false);
					rdbtnSetSkillRating.setVisible(false);
				} else {
					lblImportSheet.setVisible(true);
					textArea.setVisible(true);
					btnBrowse.setVisible(true);
					rdbtnResetSkillRating.setVisible(true);
					rdbtnSetSkillRating.setVisible(true);
				}
			}
		});
		chckbxResetAllBoards.setBounds(129, 78, 261, 23);
		contentPane.add(chckbxResetAllBoards);

		// Drag and Drop of Excel file in the Text Field
		textArea.setDropTarget(new DropTarget() {
			public synchronized void drop(DropTargetDropEvent evt) {
				try {
					evt.acceptDrop(DnDConstants.ACTION_COPY);
					@SuppressWarnings("unchecked")
					List<File> droppedFiles = (List<File>) evt.getTransferable()
							.getTransferData(DataFlavor.javaFileListFlavor);
					for (File file : droppedFiles) {
						textArea.setText(file.getAbsolutePath());
					}
				} catch (Exception ex) {
					ex.printStackTrace();
				}
			}
		});

		JButton btnStart = new JButton("START");
		btnStart.setForeground(new Color(0, 100, 0));
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// Storing Excel File location

				String excelFile = textArea.getText();

				// Getting index of the Selected Sandbox
				int sandboxItem = comboBox.getSelectedIndex();

				// Initializing Web Driver
				WebDriverManager.chromedriver().setup();
				WebDriver driver = new ChromeDriver();

				// Initializing an Explicit wait
				WebDriverWait wait = new WebDriverWait(driver, 120);

				// Opening TGO Portal
				driver.get("https://tgoportal.ubisoft.org/#/uat/homepage");
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				driver.findElement(By.xpath("/html/body/div[8]/div/div[5]/a[1]")).click();
				try {
					Thread.sleep(1000);
				} catch (InterruptedException e3) {
					e3.printStackTrace();
				}

				// Login
				driver.findElement(By.xpath("//*[@id=\"login-btn\"]")).click();
				driver.switchTo().frame("admin-connect-login-frame-0");
				try {
					Thread.sleep(500);
				} catch (InterruptedException e3) {
					e3.printStackTrace();
				}
				try {
				driver.findElement(By.xpath("//*[@id=\"login-form\"]/form/div[1]/span[2]/input")).sendKeys(username);
				driver.findElement(By.xpath("//*[@id=\"login-form\"]/form/div[3]/span[2]/input")).sendKeys(password);
				driver.findElement(By.xpath("//*[@id=\"login-form\"]/form/div[9]/input")).click();
				} catch (org.openqa.selenium.NoSuchElementException e1) {
					JOptionPane.showMessageDialog(null, "Username or Password incorrect");
					driver.quit();
				}

				// Getting handle of the 1st Window
				String handle1 = driver.getWindowHandle();

				Actions action = new Actions(driver);

				// Opening Sandbox Page
				action.moveToElement(driver.findElement(By.xpath("//*[@id=\"leftNav-games-help\"]/a/span"))).perform();
				action.moveToElement(
						driver.findElement(By.xpath("//*[@id=\"leftNav-games-help\"]/div/div/div[4]/div[3]/div[1]/a")))
						.click().perform();

				try {
					Thread.sleep(1000);
				} catch (InterruptedException e1) {
					e1.printStackTrace();
				}

				driver.switchTo().frame(0);

				// Storing the Web Element for Sandbox Search
				WebElement eSandboxSearch = driver.findElement(By.xpath(
						"/html/body/div[3]/div[2]/portal-navigation/div/sandboxes-page/sandboxes-list/sandboxes-filters/sandboxes-page-body/div/h2/div/div[2]/input"));

				// Selecting Sandbox
				String key = null;
				if (sandboxItem == 0) {
					key = "HERO_PC_UAT_AA";
					eSandboxSearch.sendKeys(key);
				} else if (sandboxItem == 1) {
					key = "HERO_PC_UAT_T";
					eSandboxSearch.sendKeys(key);
				} else if (sandboxItem == 2) {
					eSandboxSearch.sendKeys("HERO_PC_UAT_F");
				} else if (sandboxItem == 3) {
					key = "HERO_XBOXONE_UAT_O";
					eSandboxSearch.sendKeys(key);
				} else if (sandboxItem == 4) {
					key = "HERO_XBOXONE_UAT_D";
					eSandboxSearch.sendKeys(key);
				} else if (sandboxItem == 5) {
					key = "HERO_PS4_UAT_C";
					eSandboxSearch.sendKeys(key);
				} else if (sandboxItem == 6) {
					key = "HERO_PS4_UAT_M";
					eSandboxSearch.sendKeys(key);
				}

				// Clicking on the Sandbox
				try {
					wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandboxes-page/sandboxes-list/sandboxes-filters/sandboxes-page-body/div/div[2]/sandboxes-table/sandboxes-table-layout-default/table/tbody/tr/td/sandbox-detail/sandboxes-table-row/sandboxes-table-layout-default-row/div[1]/div[1]/a"))))
							.click();
				} catch (org.openqa.selenium.NoSuchElementException e1) {
					try {
						Thread.sleep(1000);
					} catch (InterruptedException e2) {
						e2.printStackTrace();
					}
					wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandboxes-page/sandboxes-list/sandboxes-filters/sandboxes-page-body/div/div[2]/sandboxes-table/sandboxes-table-layout-default/table/tbody/tr/td/sandbox-detail/sandboxes-table-row/sandboxes-table-layout-default-row/div[1]/div[1]/a"))))
							.click();
				}
				try {
					Thread.sleep(4000);
				} catch (InterruptedException e1) {
					e1.printStackTrace();
				}
				JavascriptExecutor js = (JavascriptExecutor) driver;

				// Selecting Skill Rating in the page
				try {
					// Scrolling to the bottom of the page
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
					// Clicking on skill rating
					driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandbox-page/sandbox-detail/sandbox-status/sandbox-validation/sandbox-configurations/sandbox-nodes-list/sandbox-userscripts/sandbox-navigation/sandbox-page-body/div[2]/sandbox-overview-body/div/div[2]/div[2]/div/div/div[29]/sandbox-service/div/div/div[1]/a"))
							.click();
				} catch (org.openqa.selenium.ElementClickInterceptedException e1) {
					try {
						Thread.sleep(1000);
					} catch (InterruptedException e2) {
						e2.printStackTrace();
					}
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
					driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandbox-page/sandbox-detail/sandbox-status/sandbox-validation/sandbox-configurations/sandbox-nodes-list/sandbox-userscripts/sandbox-navigation/sandbox-page-body/div[2]/sandbox-overview-body/div/div[2]/div[2]/div/div/div[29]/sandbox-service/div/div/div[1]/a"))
							.click();
				}

				// Clicking 'Rendez-Vous Management Console'
				try {
					wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandbox-page/sandbox-detail/sandbox-status/sandbox-validation/sandbox-configurations/sandbox-nodes-list/sandbox-userscripts/sandbox-navigation/sandbox-page-body/div[2]/sandbox-services-body/div[2]/div/div/div/div/div[1]/p[2]/a"))))
							.click();
				} catch (org.openqa.selenium.ElementClickInterceptedException e1) {
					js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
					wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(
							"/html/body/div[3]/div[2]/portal-navigation/div/sandbox-page/sandbox-detail/sandbox-status/sandbox-validation/sandbox-configurations/sandbox-nodes-list/sandbox-userscripts/sandbox-navigation/sandbox-page-body/div[2]/sandbox-services-body/div[2]/div/div/div/div/div[1]/p[2]/a"))))
							.click();
				}
				// Switching to the second tab
				for (String handle : driver.getWindowHandles()) {
					driver.switchTo().window(handle);
				}
				// Clicking 'Plugin - Skill'
				driver.findElement(By.xpath("/html/body/div[4]/ul[5]/li[5]/a")).click();

				// Storing the handle of second window
				String handle2 = driver.getWindowHandle();

				if (chckbxResetAllBoards.isSelected()) {
					driver.findElement(By.xpath("/html[1]/body[1]/div[4]/fieldset[2]/form[1]/p[1]/label[1]/input[1]"))
							.click();
					driver.findElement(By.xpath("/html[1]/body[1]/div[4]/fieldset[2]/form[1]/p[2]/input[1]")).click();
					try {
						driver.findElement(By.xpath("/html[1]/body[1]/div[5]/div[2]/p[1]/input[1]"))
								.sendKeys(JOptionPane.showInputDialog("Enter the key to Reset the Complete Sandbox"));
					} catch (java.lang.IllegalArgumentException e1) {
						driver.quit();
						JOptionPane.showMessageDialog(null, "The operation was cancelled or Wrong Key Entered");
					}
					driver.findElement(By.xpath("/html[1]/body[1]/div[5]/div[3]/div[1]/button[1]")).click();
					try {
						Thread.sleep(2000);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
					driver.quit();
					JOptionPane.showMessageDialog(null, "All Boards in the Database are successfully RESET");
				} else {

					// Selecting Profile page in the 1st Window
					driver.switchTo().window(handle1);
					action.moveToElement(driver.findElement(By.xpath("//*[@id=\"leftNav-players-help\"]/a/span/div")))
							.perform();
					action.moveToElement(driver
							.findElement(By.xpath("//*[@id=\"leftNav-players-help\"]/div/div/div[1]/div[1]/div[1]/a")))
							.click().perform();

					// Getting the Excel file from user location
					File src = new File(excelFile);
					FileInputStream fis = null;
					try {
						fis = new FileInputStream(src);
					} catch (FileNotFoundException e1) {
						e1.printStackTrace();
					}
					XSSFWorkbook workbook = null;
					try {
						workbook = new XSSFWorkbook(fis);
					} catch (IOException e1) {
						e1.printStackTrace();
					}

					// Getting the Excel file at the !st sheet
					XSSFSheet sheet = workbook.getSheetAt(0);

					// Storing the Last Row Number in the Excel sheet
					int rowcount = sheet.getLastRowNum();

					// Storing the boolean value of the Reset skill Radio button
					boolean SR = rdbtnResetSkillRating.isSelected();

					// For loop for operating on different Gamer Tag's, sandboxes & performing
					// SET/RESET operation
					for (int i = 2; i <= rowcount; i++) {

						// Storing the GT from excel sheet in a variable
						String Data = sheet.getRow(i).getCell(0).getStringCellValue();

						// Sending the GT to the 'Player Search' field
						action.moveToElement(driver.findElement(By.xpath(
								"//div[@id='playerPicker']//div[@class='picker-selection pseudo-input ng-pristine ng-untouched ng-valid ng-not-empty ng-valid-min-dictionary-length']")))
								.click().sendKeys(Data).perform();

						// Clicking on XBox Live if Sandbox Selected is XONE Sandbox
						if (sandboxItem == 3 || sandboxItem == 4) {
							driver.findElement(By.xpath(
									"//player-picker[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/select[1]/option[3]"))
									.click();
						}

						// Clicking on PSN if Sandbox Selected is PS4 Sandbox
						else if (sandboxItem == 5 || sandboxItem == 6) {
							driver.findElement(By.xpath(
									"//player-picker[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/select[1]/option[4]"))
									.click();
						}

						try {
							Thread.sleep(2000);
						} catch (InterruptedException e2) {
							e2.printStackTrace();
						}

						// Clicking on the Player GT from the drop down on the 'Player Search' field
						driver.findElement(By.xpath(
								"//*[@id=\"game-pickers\"]/div[3]/div[2]/div/ul/li[1]/span/span[3]/label/span[1]"))
								.click();

						try {
							Thread.sleep(3000);
						} catch (InterruptedException e2) {
							e2.printStackTrace();
						}

						// (For Copying XONE Profile ID or Copying PSN profile ID) Clicking on Copy to
						// Clipboard button if the Sandbox Selected is XONE or PS4 Sandbox
						if (sandboxItem == 3 || sandboxItem == 4 || sandboxItem == 5 || sandboxItem == 6) {
							driver.findElement(By.xpath(
									"//*[@id=\"content\"]/form/div/div[2]/linked-profile/div/form/div/table/tbody/tr[1]/td[2]/copy-text-button/a"))
									.click();
						}

						// (For Copying Uplay Profile ID) Clicking on Copy to Clipboard button if the
						// Sandbox Selected is PC Sandbox
						else {
							driver.findElement(By.xpath(
									"//*[@id=\"content\"]/form/div/div[1]/linked-profile/div/form/div/table/tbody/tr[1]/td[2]/copy-text-button"))
									.click();
						}

						// Switching to the Second tab
						driver.switchTo().window(handle2);

						// Initializing the toolkit and Clipboard for getting the copied Profile ID's
						Toolkit toolkit = Toolkit.getDefaultToolkit();
						Clipboard clipboard = toolkit.getSystemClipboard();
						String profile = null;
						try {
							profile = (String) clipboard.getData(DataFlavor.stringFlavor);
						} catch (UnsupportedFlavorException | IOException e2) {
							e2.printStackTrace();
						}

						// Sending the Profile ID to the text field in Rendez-Vous Management Console.
						driver.findElement(By.xpath("/html/body/div[4]/fieldset[1]/div/form/div/input"))
								.sendKeys(profile);

						// Clicking on Search
						driver.findElement(By.xpath("/html/body/div[4]/fieldset[1]/div/form/input")).click();

						// Checking if Radio Button 'RESET' is selected
						if (SR == true) {
							// Selecting the checkbox
							driver.findElement(
									By.xpath("//*[@id=\"frm_reset_selected\"]/table/tbody/tr[1]/th[1]/input")).click();

							// Clicking 'Reset Player' Button
							driver.findElement(By.xpath("//*[@id=\"action\"]")).click();

							// Accepting the alert
							driver.switchTo().alert().accept();

							try {
								Thread.sleep(2000);
							} catch (InterruptedException e1) {
								e1.printStackTrace();
							}

							// Clearing the Textfield for inputing next profile ID
							driver.findElement(By.xpath("/html/body/div[4]/fieldset[1]/div/form/div/input")).clear();
						}

						// Performing 'Set Skill' Operation if the Radio button 'Set' is selected
						else {
							// Clicking on the Edit Button
							driver.findElement(By.xpath("//*[@id=\"edit\"]")).click();

							// Getting the value of the Duel' skill rating from the excel sheet and sending
							// the value to the respective field after clearing the field
							try {
								int DataDuel = (int) sheet.getRow(i).getCell(1).getNumericCellValue();
								WebElement b1 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[2]/td[4]/input[1]"));
								b1.clear();
								b1.sendKeys(String.valueOf(DataDuel));
							}
							// If no data for 'Duel' is present in Excel sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Getting the value of the 'Ranked Duel' skill rating from the excel sheet and
							// sending the value to the respective field after clearing the field
							try {
								int DataRDuel = (int) sheet.getRow(i).getCell(2).getNumericCellValue();
								WebElement b2 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[5]/td[4]/input[1]"));
								b2.clear();
								b2.sendKeys(String.valueOf(DataRDuel));
							}
							// If no data for 'Ranked Duel' is present in Excel Sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Getting the value of the 'Kill' skill rating from the excel sheet and sending
							// the value to the respective field after clearing the field
							try {
								int DataKill = (int) sheet.getRow(i).getCell(3).getNumericCellValue();
								WebElement b3 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[8]/td[4]/input[1]"));
								b3.clear();
								b3.sendKeys(String.valueOf(DataKill));
							}
							// If no data for 'Kill' is present in Excel Sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Getting the value of the 'Brawl' skill rating from the excel sheet and
							// sending the value to the respective field after clearing the field
							try {
								int DataBrawl = (int) sheet.getRow(i).getCell(4).getNumericCellValue();
								WebElement b4 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[9]/td[4]/input[1]"));
								b4.clear();
								b4.sendKeys(String.valueOf(DataBrawl));
							}
							// If no data for Brawl is present in Excel Sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Getting the value of the 'Objective' skill rating from the excel sheet and
							// sending the value to the respective field after clearing the field
							try {
								int DataObjective = (int) sheet.getRow(i).getCell(5).getNumericCellValue();
								WebElement b5 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[10]/td[4]/input[1]"));
								b5.clear();
								b5.sendKeys(String.valueOf(DataObjective));
							}
							// If no data for 'Objective' is present in Excel Sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Getting the value of the 'Breach' skill rating from the excel sheet and
							// sending the value to the respective field after clearing the field
							try {
								int DataBreach = (int) sheet.getRow(i).getCell(6).getNumericCellValue();
								WebElement b6 = driver.findElement(By.xpath(
										"/html[1]/body[1]/div[4]/fieldset[1]/form[1]/table[1]/tbody[1]/tr[11]/td[4]/input[1]"));
								b6.clear();
								b6.sendKeys(String.valueOf(DataBreach));
							}
							// If no data for 'Breach' is present in Excel sheet
							catch (java.lang.NullPointerException e1) {
								e1.printStackTrace();
							}

							// Clicking the 'Update' button
							driver.findElement(By.xpath("/html[1]/body[1]/div[4]/fieldset[1]/form[1]/p[1]/input[2]"))
									.click();

							try {
								Thread.sleep(3000);
							} catch (InterruptedException e1) {
								e1.printStackTrace();
							}

							// Clicking 'Go back to Search'
							driver.findElement(By.xpath("/html[1]/body[1]/div[4]/fieldset[1]/form[1]/p[1]/a[1]"))
									.click();
						}
						// Switching to Tab 1
						driver.switchTo().window(handle1);
					}
					// Closing the Excel Sheet
					try {
						workbook.close();
					} catch (IOException e1) {
						e1.printStackTrace();
					}
					// Closing the browser
					driver.quit();

					// Show Message Dialog on Completion
					if (SR == true) {
						JOptionPane.showMessageDialog(null, "SKILL RESET SUCCESSFULLY");
					} else
						JOptionPane.showMessageDialog(null, "SKILL SET SUCCESSFULLY");
				}
			}
		});
		btnStart.setFont(new Font("Times New Roman", Font.BOLD, 15));
		btnStart.setBounds(30, 183, 221, 52);
		contentPane.add(btnStart);

		JLabel lblWelcome = new JLabel("Hello!! ");
		lblWelcome.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblWelcome.setHorizontalAlignment(SwingConstants.TRAILING);
		lblWelcome.setBounds(25, 11, 129, 29);
		contentPane.add(lblWelcome);

	}
}

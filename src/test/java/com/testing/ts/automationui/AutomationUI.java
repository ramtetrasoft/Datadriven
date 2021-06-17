package com.testing.ts.automationui;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;
import javax.swing.UIManager;

@SuppressWarnings("serial")
public class AutomationUI extends javax.swing.JFrame
{
	public AutomationUI() {
        initComponents();
    }
	 private void initComponents() {
	        jLabel1 = new javax.swing.JLabel();
	        runSanitySuite = new javax.swing.JButton();
	        viewReports = new javax.swing.JButton();
	        editViewTestData = new javax.swing.JButton();
	        config = new javax.swing.JButton();
	        jLabel2 = new javax.swing.JLabel();
	        jLabel3 = new javax.swing.JLabel();

	        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
	        setTitle(" Automation Sanity Run");

	        jLabel1.setText("       BI AUTOMATION SANITY RUN");

	        runSanitySuite.setText("Run Sanity Suite");
	        runSanitySuite.setToolTipText("Click to start execution of automated suite");
	        runSanitySuite.addActionListener(new java.awt.event.ActionListener() {
	            public void actionPerformed(java.awt.event.ActionEvent evt) {
	                runSanitySuiteActionPerformed(evt);
	            }
	        });

	        viewReports.setText("View Reports");
	        viewReports.setToolTipText("Click to view the ATU reports after execution");
	        viewReports.addActionListener(new java.awt.event.ActionListener() {
	            public void actionPerformed(java.awt.event.ActionEvent evt) {
	                viewReportsActionPerformed(evt);
	            }
	        });
	        
	        editViewTestData.setText("Edit/View Test Scenarios and Data");
	        editViewTestData.setToolTipText("Click to view/change the test data and the scripts to be executed before execution");
	        editViewTestData.addActionListener(new java.awt.event.ActionListener() {
	            public void actionPerformed(java.awt.event.ActionEvent evt) {
	                editViewTestDataActionPerformed(evt);
	            }
	        });

	        config.setText("Configure URL,Credentials & Browser");
	        config.setToolTipText("Change the URL/Login Credentials & Browser");
	        config.addActionListener(new java.awt.event.ActionListener() {
	            public void actionPerformed(java.awt.event.ActionEvent evt) {
	                configActionPerformed(evt);
	            }
	        });

	        jLabel2.setText("Note: Kindly close the test data excel sheet while starting the utility");
	        jLabel3.setText("Do not disturb the system while running the utility");

	        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
	        getContentPane().setLayout(layout);
	        layout.setHorizontalGroup(
	            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
	            .addGroup(layout.createSequentialGroup()
	                .addGap(33, 33, 33)
	                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
	                    .addComponent(jLabel3)
	                    .addComponent(jLabel2))
	                .addContainerGap(43, Short.MAX_VALUE))
	            .addGroup(layout.createSequentialGroup()
	                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
	                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
	                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 199, javax.swing.GroupLayout.PREFERRED_SIZE)
	                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
	                        .addComponent(runSanitySuite, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
	                        .addComponent(viewReports, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
	                        .addComponent(editViewTestData, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
	                        .addComponent(config, javax.swing.GroupLayout.DEFAULT_SIZE, 208, Short.MAX_VALUE)))
	                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
	        );
	        layout.setVerticalGroup(
	            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
	            .addGroup(layout.createSequentialGroup()
	                .addContainerGap()
	                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
	                .addGap(18, 18, 18)
	                .addComponent(config)
	                .addGap(18, 18, 18)
	                .addComponent(editViewTestData)
	                .addGap(18, 18, 18)
	                .addComponent(runSanitySuite)
	                .addGap(18, 18, 18)
	                .addComponent(viewReports)
	                .addGap(33, 33, 33)
	                .addComponent(jLabel2)
	                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
	                .addComponent(jLabel3)
	                .addContainerGap(27, Short.MAX_VALUE))
	        );

	        pack();
	    }	private void runSanitySuiteActionPerformed(java.awt.event.ActionEvent evt) {
			try {
				/*System.out.println(System.getProperty("user.dir"));
				Runtime.getRuntime().exec("cmd /c start \"\" "+ fileLocation + batchName);*/
				String batch= "cmd /c start \"\""+" "+"\""+ System.getProperty("user.dir")+System.getProperty("file.separator")+"\\seleniumautots.bat";
				//String batch= "cmd /c start \"\""+" "+"\""+"C:\\Users\\tsipl1805\\eclipse-workspace\\org.ts.seleniumauto\\SeleniumAutoTS_Sanity.bat";
				Runtime.getRuntime().exec(batch);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	    public void infoBox(String infoMessage, String titleBar)
	    {
	        infoBox(infoMessage, titleBar, null);
	    }
	    public void infoBox(String infoMessage, String titleBar, String headerMessage)
	    {
	        JOptionPane.showMessageDialog(null, infoMessage, "InfoBox: " + titleBar, JOptionPane.INFORMATION_MESSAGE);
	    }
		private void viewReportsActionPerformed(java.awt.event.ActionEvent evt) {
			File htmlFile = new File(System.getProperty("user.dir")+System.getProperty("file.separator") + "\\ATU Reports\\index.html");
			try {
				Desktop.getDesktop().browse(htmlFile.toURI());
			} catch (IOException e) {
				infoBox("File Location Not Found!!!","InfoBox");
			}
		}

		private void editViewTestDataActionPerformed(java.awt.event.ActionEvent evt) {
			String filePath = System.getProperty("user.dir")+System.getProperty("file.separator") + "\\Testdata\\DataSheet.xls";
			try {
				if (Desktop.isDesktopSupported()) {
					Desktop.getDesktop().open(new File(filePath));
				} else {
					infoBox("File Location Not Found!!!","InfoBox");
				}

			} catch (IOException e) {
				infoBox("File Location Not Found!!!","InfoBox");
			}
		}

		private void configActionPerformed(java.awt.event.ActionEvent evt) {
			String filePath = System.getProperty("user.dir")+System.getProperty("file.separator") + "\\Properties\\config.properties";
			try {
				if (Desktop.isDesktopSupported()) {
					Desktop.getDesktop().open(new File(filePath));
				} else {
					infoBox("File Location Not Found!!!","InfoBox");
				}

			} catch (IOException e) {
				infoBox("File Location Not Found!!!","InfoBox");
			}
		}
		
	    public static void main(String args[]) {
	       try {
	    	   UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
	        } catch (ClassNotFoundException ex) {
	            java.util.logging.Logger.getLogger(AutomationUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
	        } catch (InstantiationException ex) {
	            java.util.logging.Logger.getLogger(AutomationUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
	        } catch (IllegalAccessException ex) {
	            java.util.logging.Logger.getLogger(AutomationUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
	        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
	            java.util.logging.Logger.getLogger(AutomationUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
	        }
	       

	        /* Create and display the form */
	        java.awt.EventQueue.invokeLater(new Runnable() {
	            public void run() {
	                new AutomationUI().setVisible(true);
	            }
	        });
	    }

	    // Variables declaration - do not modify                     
	    private javax.swing.JButton config;
	    private javax.swing.JButton editViewTestData;
	    private javax.swing.JLabel jLabel1;
	    private javax.swing.JLabel jLabel2;
	    private javax.swing.JLabel jLabel3;
	    private javax.swing.JButton runSanitySuite;
	    private javax.swing.JButton viewReports;
	      // End of variables declaration
	}

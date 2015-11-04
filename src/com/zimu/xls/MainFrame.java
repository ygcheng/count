package com.zimu.xls;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.EventQueue;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Collections;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;

import createxls.ExcelHandle;
import createxls.Status;
import createxls.TimeBucket;

public class MainFrame extends JFrame implements ActionListener{
	
	private JPanel contentPane;
	private JTextArea text_controlText;
	private JTextField text_feng;
	private JTextField text_ping;
	private JTextField text_gu;
	private JTextField text_filePath;
	private JButton btn_fileChooser;
	private JButton btn_reset;
	private JButton btn_run;
	private final Font font = new Font("Default",Font.PLAIN,16);
	
	public MainFrame(){
		initView();
	}

	private void initView() {
		
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setBounds(100, 100, 900, 600);
		this.setTitle("节能统计");
		this.contentPane = new JPanel();
		this.contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		this.contentPane.setLayout(new BorderLayout());
		
		JPanel north = new JPanel();
		north.setLayout(new GridLayout(3, 1));
		
		JPanel north1 = new JPanel();
		JPanel north2 = new JPanel();
		JPanel north3 = new JPanel();
		
		north1.setLayout(new FlowLayout(FlowLayout.LEADING));
		this.text_filePath = new JTextField(50);
		this.text_filePath.setEditable(false);
		north1.add(this.text_filePath);
		this.btn_fileChooser = new JButton(" 选择文件 ");
		this.btn_fileChooser.addActionListener(this);
		north1.add(this.btn_fileChooser);
		
		north2.setLayout(new FlowLayout(FlowLayout.LEADING));
		JLabel label1 = new JLabel("峰（元/小时）：");
		north2.add(label1);
		this.text_feng = new JTextField(10);
		north2.add(this.text_feng);
		JLabel label2 = new JLabel("平（元/小时）：");
		north2.add(label2);
		this.text_ping = new JTextField(10);
		north2.add(this.text_ping);
		JLabel label3 = new JLabel("谷（元/小时）：");
		north2.add(label3);
		this.text_gu = new JTextField(10);
		north2.add(this.text_gu);
		
		north3.setLayout(new FlowLayout(FlowLayout.TRAILING));
		this.btn_reset = new JButton(" 重置 ");
		this.btn_reset.addActionListener(this);
		north3.add(this.btn_reset);
		this.btn_run = new JButton(" 开始运行 ");
		this.btn_run.addActionListener(this);
		north3.add(this.btn_run);
		
		north.add(north1);
		north.add(north2);
		north.add(north3);
		
		this.text_controlText = new JTextArea();
		this.text_controlText.setEditable(false);
		this.text_controlText.setBorder(new EmptyBorder(5, 5, 5, 5));
		this.text_controlText.setBackground(new Color(169, 169, 169));
		JScrollPane scroll = new JScrollPane(this.text_controlText);
		//分别设置水平和垂直滚动条自动出现 
		scroll.setHorizontalScrollBarPolicy( 
		JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED); 
		scroll.setVerticalScrollBarPolicy( 
		JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		
		this.contentPane.add(BorderLayout.NORTH,north);
		this.contentPane.add(BorderLayout.CENTER,scroll);
		this.setContentPane(contentPane);
		
		
		this.text_controlText.setFont(font);
		label1.setFont(font);
		this.text_feng.setFont(font);
		label2.setFont(font);
		this.text_ping.setFont(font);
		label3.setFont(font);
		this.text_gu.setFont(font);
		this.text_filePath.setFont(font);
		this.btn_fileChooser.setFont(font);
		this.btn_reset.setFont(font);
		this.btn_run.setFont(font);
	}

	public static void main(String[] args) {
		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainFrame frame = new MainFrame();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
		
	}

	@Override
	public void actionPerformed(ActionEvent e) {

			if(e.getSource()==this.btn_fileChooser){
				
				JFileChooser fDialog = new JFileChooser();
				fDialog.setFont(font);
				fDialog.showOpenDialog(this);
				File file = fDialog.getSelectedFile();
				this.text_filePath.setText(file.getPath());
				
			}else if(e.getSource()==this.btn_run){
				this.text_controlText.setText(null);
				
				String filePath = this.text_filePath.getText();
				String strfeng = this.text_feng.getText();
				String strping = this.text_ping.getText();
				String strgu = this.text_gu.getText();
				
				if(filePath == null||"".equals(filePath)){
					this.text_controlText.append("请选择文件" + "\n");
					return;
				}
				
				float feng;
				float ping;
				float gu;
				try{
					feng = Float.parseFloat(strfeng);
					ping = Float.parseFloat(strping);
					gu = Float.parseFloat(strgu);
				}catch(Exception exception){
					this.text_controlText.append("价格输入部分有问题" + "\n");
					System.out.println(exception.getCause().getMessage());
					return;
				}
				
				this.text_controlText.append("文件路径:" + filePath + "\n");
				this.text_controlText.append("峰（元/小时）:" + feng + "\n");
				this.text_controlText.append("平（元/小时）:" + ping + "\n");
				this.text_controlText.append("谷（元/小时）:" + gu + "\n");
				
				try {
					this.count(filePath, feng, ping, gu);
				} catch (Exception e1) {
					e1.printStackTrace();
					this.text_controlText.append(e1.getCause().getMessage() + "\n");
				}
			}else if(e.getSource()==this.btn_reset){
				this.text_filePath.setText(null);
				this.text_feng.setText(null);
				this.text_ping.setText(null);
				this.text_gu.setText(null);
				this.text_controlText.setText(null);
			}
		
	}
	
	
	private void count(String filePath,float feng,float ping,float gu) throws Exception{
		// 读EXCEL 生成对象原始数据
		List<Status> original = ExcelHandle.readExcel(filePath);
		//List<Status> original = ExcelHandle.readExcel("E:/文档/Documents/舒浩哥/6月二期.xls");
		
		if(original == null){
			this.text_controlText.append("读EXCEL失败" + "\n");
			return;
		}
				
		//把原始数据根据机柜名称分组
		List<List<Status>> group = ExcelHandle.originalGrouping(original);
		this.text_controlText.append("\n");
		this.text_controlText.append("原始数据分组 共" + group.size() + "组" + "\n");
					
		//选出运行时间段
		List<List<TimeBucket>> groupTB = ExcelHandle.originalTB(group);
		for(List<TimeBucket> list : groupTB){
			float t1 = 0f;
			if(list.size()>0){
				System.out.println(list.get(0).getStart().getName());
				this.text_controlText.append(list.get(0).getStart().getName() +"\n");
			}
			for(TimeBucket tb : list){
				System.out.println(tb.toString());
				this.text_controlText.append(tb.toString() +"\n");
				t1 +=  ((tb.getEnd().getTime().getTime()-tb.getStart().getTime().getTime())/3600000f);
			}
			System.out.println("共运行了 ： " + t1 + " 个小时");
			this.text_controlText.append("共运行了 ： " + t1 + " 个小时" +"\n");
			this.text_controlText.append("\n");
		}
		
		
		//选出运行时间段（按峰平谷把时间段割小）
		List<List<TimeBucket>> groupTB2 = ExcelHandle.splitTB(groupTB);
		for(List<TimeBucket> list : groupTB2){
			Collections.sort(list);
		}

		
		for(List<TimeBucket> list : groupTB2){
			float t1 = 0f;
			float t2 = 0f;
			float t3 = 0f;
			for(TimeBucket tb : list){
				//每个小时的时间区间差值3600000
				float hour = (tb.getEnd().getTime().getTime()-tb.getStart().getTime().getTime())/3600000.0f;
				System.out.println(tb.toString());
				this.text_controlText.append(tb.toString() +"\n");
				if(tb.getTrend() == 1){
					t1 += hour;
				}else if(tb.getTrend() == 2){
					t2 += hour;
				}else if(tb.getTrend() == 3){
					t3 += hour;
				}
			}
			if(list.size() > 0){
				System.out.println("组：" + list.get(0).getEnd().getName());
				System.out.println("峰：" + t1 + "小时");
				System.out.println("平：" + t2 + "小时");
				System.out.println("谷：" + t3 + "小时");
				this.text_controlText.append("组：" + list.get(0).getEnd().getName()+"\n");
				this.text_controlText.append("峰：" + t1 + "小时"+"\n");
				this.text_controlText.append("平：" + t2 + "小时"+"\n");
				this.text_controlText.append("谷：" + t3 + "小时"+"\n");
			}
			
		}
		for(int i = 0; i < 10 ;i++){
			this.text_controlText.append("最后结果--------------------------------\n");
		}
		for(List<TimeBucket> list : groupTB2){
			float t1 = 0f;
			float t2 = 0f;
			float t3 = 0f;
			for(TimeBucket tb : list){
				//每个小时的时间区间差值3600000
				float hour = (tb.getEnd().getTime().getTime()-tb.getStart().getTime().getTime())/3600000.0f;
				if(tb.getTrend() == 1){
					t1 += hour;
				}else if(tb.getTrend() == 2){
					t2 += hour;
				}else if(tb.getTrend() == 3){
					t3 += hour;
				}
			}
			if(list.size() > 0){
				this.text_controlText.append("\n");
				this.text_controlText.append("组：" + list.get(0).getEnd().getName() + "\n");
				this.text_controlText.append("峰：" + t1 + "小时 - 计" + t1*feng + "元" + "\n");
				this.text_controlText.append("平：" + t2 + "小时 - 计" + t2*ping + "元" + "\n");
				this.text_controlText.append("谷：" + t3 + "小时 - 计" + t3*gu + "元" + "\n");
			}
			
		}
		
	}
	
	
	
	
	
}

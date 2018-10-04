<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Welcome extends CI_Controller {

	public function index(){
		$this->load->library("excel");
		$this->excel->createSheet();
		$this->excel->setActiveSheetIndex(0);
		$label = array(1=>'A',2=>'B',3=>'C',4=>'D',5=>'E',6=>'F',7=>'G',8=>'H',9=>'I',10=>'J',11=>'K',12=>'L',13=>'M',14=>'N',15=>'O',16=>'P',17=>'Q',18=>'R',19=>'S',20=>'T',21=>'U',22=>'V',23=>'W',24=>'X',25=>'Y',26=>'Z');
		$column = 1;
		$row = 1;
		foreach ($lists[0] as $key => $value) {
			$this->excel->getActiveSheet()->SetCellValue($label[$column].$row, $key);
			$column = $column+1;
		}
		$row = 2;
		foreach ($lists as $key => $list) {
			$column = 1;
			foreach ($list as $key => $value) {
				$this->excel->getActiveSheet()->SetCellValue($label[$column].$row, $value);
				$column = $column+1;
			}
			$row = $row+1;
		}  
		$this->excel->stream('Weaving_cost_'.date('Y_m_d').'.xls');
	}

}

/* End of file welcome.php */
/* Location: ./application/controller/welcome.php */
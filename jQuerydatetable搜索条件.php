<?php 
 
    /**
     * 获取数据分页======单表查询列表
     */
    public function actionAjaxgetdata(){
        $models = $this->_userModel->find();
        //获取post集合
        $postData = $this->_post();
        $this->setSession('order_post_data' , $postData);
        //循环表格列，获取有搜索值得列拼接where sql
        foreach ($postData['columns'] as $columns){
            $searchVal = trim($columns['search']['value']);
            if($searchVal){
                //值区间搜索
                if(strstr($searchVal , '-yadcf_delim-')){
                    $interVal = explode('-yadcf_delim-' , $searchVal);
                    if($interVal[0]){
                        $models->andWhere(['>=', $columns['data'], $interVal[0]]);
                    }
                    if($interVal[1]){
                        $models->andWhere(['<=', $columns['data'], $interVal[1]]);
                    }
                }else{  //值搜索
                    
                    $models->andWhere(['LIKE', $columns['data'], trim($columns['search']['value'])]);
                }
            }
        }
        if(isset($postData['search']['value']) && trim($postData['search']['value'])){
            foreach ($postData['columns'] as $columns){
                if(!is_numeric($columns['data'])){
                    $models->orWhere(['LIKE', $columns['data'], trim($postData['search']['value'])]);
                }
            }
        }
        //查询排序
        $orderBy = $postData['columns'][$postData['order'][0]['column']]['data'] . " " .$postData['order'][0]['dir'];
        $models->orderBy($orderBy);
        //分页查询
        $orderList = $models->offset($postData['start'])->limit($postData['length'])->asArray()->all();
        if(!empty($orderList)){
            for($i=0;$i<count($orderList);$i++){
                $orderList[$i]['BEGINDATE'] = date('Y-m-d', $orderList[$i]['BEGINDATE']);
                $orderList[$i]['ENDDATE'] = date('Y-m-d', $orderList[$i]['ENDDATE']);
            }
        }
        $data['draw']               = Yii::$app->request->post('draw');
        $data['recordsTotal']       = $models->count();
        $data['recordsFiltered']    = $data['recordsTotal'];
        $data['data']               = $orderList;
        Response::returnJson(json_encode($data));
    }
    
    
    
    
    /**
    *
    *多表
    *
    */
        function GetCustomernData($param)
    {
        try {    
            
            $query = $this->GetCondition($param);
            
            if($query != null)
            {
                //页码
                $data['draw'] = $query['draw'];
                //总数，分页总数
                $data['recordsTotal'] = $data['recordsFiltered'] = $query['query']->from('ECMS_COMPANY A')
                    ->leftJoin('ECMS_CUSTOMER B','A.ID=B.COMPANYID')->count('A.ID');
                //分页查询条件
                $query['query']->offset($query['start'])->limit($query['length']);
                //排序
                if(!empty($query['order']))
                    $query['query']->orderBy($query['order']);
                //查询数据集
                $data['data'] = $query['query']->select("B.ID,B.TITLE as NAME,B.CODE,B.TRUENAME,B.MOBILE,B.BEGINDATE,B.ENDDATE,B.CREDIT,A.TITLE")->asArray()->all();
                //重造数据显示值
               if(!empty($data['data']))
                {
                    for ($i = 0; $i < count($data['data']); $i++)
                    {
                        $data['data'][$i]['BEGINDATE'] = date('Y-m-d', $data['data'][$i]['BEGINDATE']);
                        $data['data'][$i]['ENDDATE'] = date('Y-m-d', $data['data'][$i]['ENDDATE']);
                    }
                }
                
                return $data;
            }
        }
        catch(\Exception $ex)
        {
            if (YII_DEBUG) {
                throw new \Exception($ex->getMessage(), $ex->getCode(), $ex->getPrevious());
            }
        }
    
        return null;
    }
    
    
    
    
     function GetCondition($param)
    {       
        try {           
            if(is_array($param) && count($param) > 0)
            {
                $data['draw'] = is_numeric($param['draw']) ? $param['draw'] : 1;
        
                $query = $this->find();
                
                foreach ($param as $k => $v)
                {
                    if($k == 'columns'){
                        foreach ($v as $column){
                            $searchVal = trim($column['search']['value']);
                            if($searchVal){
                                $searchName = !empty($column['name']) ? $column['name'] : $column['data'];
                                //值区间搜索
                                if(strstr($searchVal , '-yadcf_delim-')){
                                    $interVal = explode('-yadcf_delim-' , $searchVal);
                                    if($interVal[0]){
                                        $query->andFilterWhere(['>=', $searchName, $interVal[0]]);
                                    }
                                    if($interVal[1]){
                                        $query->andFilterWhere(['<=', $searchName, $interVal[1]]);
                                    }
                                }else{  //值搜索
                                    $query->andFilterWhere(['LIKE', $searchName, trim($column['search']['value'])]);
                                }
                            }
                        }
                    }elseif(isset($v[0]) && is_array($v[0]))
                    {                        
                        foreach ($v as $k1 => $v1)
                        {
                            $query = $this->GetWhere($query,$k,$v1);
                        }
                    }
                    else
                    {
                        $query = $this->GetWhere($query,$k,$v);
                    }            
                }
                $data['start'] = is_numeric($param['start']) ? $param['start'] : 0;
                $data['length'] = is_numeric($param['length']) ? $param['length'] : 15;
                
                if(!empty($param['columns']) && !empty($param['order']))
                {
                    if(count($param['columns']) > $param['order'][0]['column']
                        && !empty($param['columns'][$param['order'][0]['column']])
                        && !empty($param['columns'][$param['order'][0]['column']]['data'])
                    )
                    {
                        $data['order'] = $param['columns'][$param['order'][0]['column']]['data'].' '.$param['order'][0]['dir'];
                    }
                }
                
                $data['query'] = $query;
                
                return $data;
            }
        
        }catch(\Exception $ex){
            if (YII_DEBUG) {
                throw new \Exception($ex->getMessage(), $ex->getCode(), $ex->getPrevious());
            }
        }
        
        return null;
    }
    /** 
     * @param array $param
     * @return array count=> data=>
     */
    function GetData($param)
    {         
        try {
            $query = $this->GetCondition($param);
            
            if($query != null)
            {
                $data['draw'] = $query['draw'];
                $data['count'] = $query['query']->count();
                
                $query['query']->offset($query['start'])->limit($query['length']);
                
                if(!empty($query['order']))
                    $query['query']->orderBy($query['order']);
                
                $data['data'] = $query['query']->all();
                //print_r($query['query']->createCommand()->getRawSql());
                return $data;
            }
        }
        catch(\Exception $ex)
        {
            if (YII_DEBUG) {
                throw new \Exception($ex->getMessage(), $ex->getCode(), $ex->getPrevious());
            }
        }
        
        return null;
    }  
    

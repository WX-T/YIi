<?php
 /*
  *
  * YII 2.0 student notes
  *  @auther WX-T
  *
  *
 */
  
  //生成Url
  // /index.php?r=site/index
  echo Url::toRoute('site/index');

  // /index.php?r=site/index&src=ref1#name
  echo Url::toRoute(['site/index', 'src' => 'ref1', '#' => 'name']);

  // /index.php?r=post/edit&id=100     assume the alias "@postEdit" is defined as "post/edit"
  echo Url::toRoute(['@postEdit', 'id' => 100]);

  // http://www.example.com/index.php?r=site/index
  echo Url::toRoute('site/index', true);

  // https://www.example.com/index.php?r=site/index
  echo Url::toRoute('site/index', 'https');
  
  //Url::to()
  // /index.php?r=site/index
  echo Url::to(['site/index']);
  
  // /index.php?r=site/index&src=ref1#name
  echo Url::to(['site/index', 'src' => 'ref1', '#' => 'name']);
  
  // /index.php?r=post/edit&id=100     assume the alias "@postEdit" is defined as "post/edit"
  echo Url::to(['@postEdit', 'id' => 100]);
  
  // the currently requested URL
  echo Url::to();
  
  // /images/logo.gif
  echo Url::to('@web/images/logo.gif');
  
  // images/logo.gif
  echo Url::to('images/logo.gif');
  
  // http://www.example.com/images/logo.gif
  echo Url::to('@web/images/logo.gif', true);
  
  // https://www.example.com/images/logo.gif
  echo Url::to('@web/images/logo.gif', 'https');
  
  /**
   *=========================================页面加载====================================================
  */
  //加载页面,同时显示头部和尾部
  $this->render('页面' , $data);
  //加载页面,不加载头部和尾部，(加载页面块，ajax弹出修改时弹出的页面)
  $this->renderPartial('html页面' , $data);
  
  //===========================================数据库查询====================================================
  
  User::->from('ECMS_ADMIN A')
  ->leftJoin('ECMS_ADMININFO B','A.ID=B.ADMINID')
  ->leftJoin('ECMS_ROLE C','C.ID=A.ROLEID')
  ->leftJoin('ECMS_DEPARTMENT D','C.DEPARTID=D.ID')->count('A.ID'); //左链接例子
  
  User::find()->all();    //此方法返回所有数据；
  
  User::findOne($id);   //此方法返回 主键 id=1  的一条数据(举个例子)； 
  
  User::find()->where(['name' => '小伙儿'])->one();   //此方法返回 ['name' => '小伙儿'] 的一条数据；
  
  User::find()->where(['name' => '小伙儿'])->all();   //此方法返回 ['name' => '小伙儿'] 的所有数据；
  
  User::find()->orderBy('id DESC')->all();   //此方法是排序查询；
  
  User::findBySql('SELECT * FROM user')->all();  //此方法是用 sql  语句查询 user 表里面的所有数据；
  
  User::findBySql('SELECT * FROM user')->one();  //此方法是用 sql  语句查询 user 表里面的一条数据；
  
  User::find()->andWhere(['sex' => '男', 'age' => '24'])->count('id');   //统计符合条件的总条数；
  
  User::find()->select(['TITLE'])->orderBy("TITLE")->asArray()->all();   //group by用法
  
  User::find()->one();    //此方法返回一条数据；
  
  User::find()->all();    //此方法返回所有数据；
  
  User::find()->count();    //此方法返回记录的数量；
  
  User::find()->average();    //此方法返回指定列的平均值；
  
  User::find()->min();    //此方法返回指定列的最小值 ；
  
  User::find()->max();    //此方法返回指定列的最大值 ；
  
  User::find()->scalar();    //此方法返回值的第一行第一列的查询结果；
  
  User::find()->column();    //此方法返回查询结果中的第一列的值；
  
  User::find()->exists();    //此方法返回一个值指示是否包含查询结果的数据行；
  
  User::find()->batch(10);  //每次取 10 条数据 
  
  User::find()->each(10);  //每次取 10 条数据， 迭代查询
//=========================读取config/params.php的配置====================
 echo Yii::$app->params['name'];
//===========================数据操作===========================
    //1.简单查询  
    $admin=Admin::model()->findAll($condition,$params);  
    $admin=Admin::model()->findAll("username=:name",array(":name"=>$username));  
       
    $infoArr= NewsList::model()->findAll("status = '1' ORDER BY id DESC limit 10 ");  
       
    //2. findAllByPk(该方法是根据主键查询一个集合,可以使用多个主键)  
    $admin=Admin::model()->findAllByPk($postIDs,$condition,$params);  
    $admin=Admin::model()->findAllByPk($id,"name like :name and age=:age",array(':name'=>$name,'age'=>$age));  
    $admin=Admin::model()->findAllByPk(array(1,2));  
       
    //3.findAllByAttributes (该方法是根据条件查询一个集合,可以是多个条件,把条件放到数组里面)  
    $admin=Admin::model()->findAllByAttributes($attributes,$condition,$params);  
    $admin=Admin::model()->findAllByAttributes(array('username'=>'admin'));  
       
    //4.findAllBySql (该方法是根据SQL语句查询一个数组)  
    $admin=Admin::model()->findAllBySql($sql,$params);  
    $admin=Admin::model()->findAllBySql("select * from admin where username like :name",array(':name'=>'%ad%'));  
    User::find()->all();    此方法返回所有数据；  
    User::findOne($id);   此方法返回 主键 id=1  的一条数据(举个例子)；   
    User::find()->where(['name' => '小伙儿'])->one();   此方法返回 ['name' => '小伙儿'] 的一条数据；  
    User::find()->where(['name' => '小伙儿'])->all();   此方法返回 ['name' => '小伙儿'] 的所有数据；  
    User::find()->orderBy('id DESC')->all();   此方法是排序查询；  
    User::findBySql('SELECT * FROM user')->all();  此方法是用 sql  语句查询 user 表里面的所有数据；  
    User::findBySql('SELECT * FROM user')->one();  此方法是用 sql  语句查询 user 表里面的一条数据；  
    User::find()->andWhere(['sex' => '男', 'age' => '24'])->count('id');   统计符合条件的总条数；  
    User::find()->one();    此方法返回一条数据；  
    User::find()->all();    此方法返回所有数据；  
    User::find()->count();    此方法返回记录的数量；  
    User::find()->average();    此方法返回指定列的平均值；  
    User::find()->min();    此方法返回指定列的最小值 ；  
    User::find()->max();    此方法返回指定列的最大值 ；  
    User::find()->scalar();    此方法返回值的第一行第一列的查询结果；  
    User::find()->column();    此方法返回查询结果中的第一列的值；  
    User::find()->exists();    此方法返回一个值指示是否包含查询结果的数据行；  
    User::find()->batch(10);  每次取 10 条数据   
    User::find()->each(10);  每次取 10 条数据， 迭代查询  
    二、查询对象的方法  
    //根据主键查询出一个对象,如：findByPk(1);  
    $admin=Admin::model()->findByPk($postID,$condition,$params);  
    $admin=Admin::model()->findByPk(1);  
       
    //根据一个条件查询出一组数据,可能是多个,但是他只返回第一行数据  
    $row=Admin::model()->find($condition,$params);  
    $row=Admin::model()->find('username=:name',array(':name'=>'admin'));  
       
    //该方法是根据条件查询一组数据,可以是多个条件,把条件放到数组里面,查询的也是第一条数据  
    $admin=Admin::model()->findByAttributes($attributes,$condition,$params);  
    $admin=Admin::model()->findByAttributes(array('username'=>'admin'));  
       
    //该方法是根据SQL语句查询一组数据,他查询的也是第一条数据  
    $admin=Admin::model()->findBySql($sql,$params);  
    $admin=Admin::model()->findBySql("select * from admin where username=:name",array(':name'=>'admin'));  
       
    //拼一个获得SQL的方法,在根据find查询出一个对象   
    $criteria=newCDbCriteria;   
    $criteria->select='username';// only select the 'title' column   
    $criteria->condition='username=:username';    //请注意,这是一个查询的条件,且只有一个查询条件.多条件用addCondition  
    $criteria->params=array(":username=>'admin'");  
    $criteria->order ="id DESC";  
    $criteria->limit ="3";  
    $post=Post::model()->find($criteria);// $params isnot needed   
       
    //多条件查询的语句  
    $criteria= new CDbCriteria;       
    $criteria->addCondition("id=1");//查询条件，即where id = 1   
    $criteria->addInCondition('id',array(1,2,3,4,5));//代表where id IN (1,2,3,4,5,);   
    $criteria->addNotInCondition('id',array(1,2,3,4,5));//与上面正好相法，是NOT IN   
    $criteria->addCondition('id=1','OR');//这是OR条件，多个条件的时候，该条件是OR而非AND   
    $criteria->addSearchCondition('name','分类');//搜索条件，其实代表了。。where name like '%分类%'   
    $criteria->addBetweenCondition('id', 1, 4);//between 1 and 4  
    $criteria->compare('id', 1);   //这个方法比较特殊,他会根据你的参数自动处理成addCondition或者addInCondition.  
    $criteria->compare('id',array(1,2,3));   //即如果第二个参数是数组就会调用addInCondition   
       
       
    $criteria->select ='id,parentid,name';//代表了要查询的字段，默认select='*';   
    $criteria->join ='xxx'; //连接表   
    $criteria->with ='xxx'; //调用relations    
    $criteria->limit = 10;   //取1条数据，如果小于0，则不作处理   
    $criteria->offset = 1;  //两条合并起来，则表示 limit 10 offset 1,或者代表了。limit 1,10   
    $criteria->order ='xxx DESC,XXX ASC' ;//排序条件   
    $criteria->group ='group 条件';   
    $criteria->having ='having 条件 ';   
    $criteria->distinct = FALSE;//是否唯一查询  
    三、查询个数,判断查询是否有结果  
    //该方法是根据一个条件查询一个集合有多少条记录,返回一个int型数字  
    $n=Post::model()->count($condition,$params);  
    $n=Post::model()->count("username=:name",array(":name"=>$username));  
       
    //该方法是根据SQL语句查询一个集合有多少条记录,返回一个int型数字  
    $n=Post::model()->countBySql($sql,$params);  
    $n=Post::model()->countBySql("select * from admin where username=:name",array(':name'=>'admin'));  
       
    //该方法是根据一个条件查询查询得到的数组有没有数据,如果有数据返回一个true,否则没有找到  
    $exists=Post::model()->exists($condition,$params);  
    $exists=Post::model()->exists("name=:name",array(":name"=>$username));  
    四、新增  
    $admin= new Admin;         
    $admin->username =$username;  
    $admin->password =$password;  
    if($admin->save() > 0){echo "添加成功"; }else{echo "添加失败"; }  
    五、修改  
    $model = Post::find()->where(['id'=>1])->one();   //查询修改的值
    $model->name = '张三'; //执行搜索
    if($model->save()){
      echo '成功';
    }else{
     echo '失败';
    }  //保存
    六、删除  
    //deleteAll  
    Post::model()->deleteAll($condition,$params);  
    $count= Admin::model()->deleteAll('username=:name and password=:pass',array(':name'=>'admin',':pass'=>'admin'));  
    $count= Admin::model()->deleteAll('id in("1,2,3")');//删除id为这些的数据  
    if($count>0){echo"删除成功"; }else{echo "删除失败"; }  
       
    //deleteByPk  
    Post::model()->deleteByPk($pk,$condition,$params);  
    $count= Admin::model()->deleteByPk(1);  
    $count=Admin::model()->deleteByPk(array(1,2),'username=:name',array(':name'=>'admin'));  
    if($count>0){echo "删除成功"; }else{echo "删除失败"; }  
 //========================直接操作数据库========================
     //createCommand(执行原生的SQL语句)  
    $sql= "SELECT u.account,i.* FROM sys_user as u left join user_info as i on u.id=i.user_id";  
    $rows=Yii::$app->db->createCommand($sql)->query();  
    foreach($rows as $k => $v){  
        echo$v['add_time'];  
    }  
      
    查询返回多行：  
      
    $command = $connection->createCommand('SELECT * FROM post');  
    $posts = $command->queryAll();  
    返回单行：  
      
    $command = $connection->createCommand('SELECT * FROM post WHERE id=1');  
    $post = $command->queryOne();  
    查询多行单值：  
      
    $command = $connection->createCommand('SELECT title FROM post');  
    $titles = $command->queryColumn();  
    查询标量值/计算值：  
      
    $command = $connection->createCommand('SELECT COUNT(*) FROM post');  
    $postCount = $command->queryScalar();  

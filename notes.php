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
  
  User::find()->all();    //此方法返回所有数据；
  
  User::findOne($id);   //此方法返回 主键 id=1  的一条数据(举个例子)； 
  
  User::find()->where(['name' => '小伙儿'])->one();   //此方法返回 ['name' => '小伙儿'] 的一条数据；
  
  User::find()->where(['name' => '小伙儿'])->all();   //此方法返回 ['name' => '小伙儿'] 的所有数据；
  
  User::find()->orderBy('id DESC')->all();   //此方法是排序查询；
  
  User::findBySql('SELECT * FROM user')->all();  //此方法是用 sql  语句查询 user 表里面的所有数据；
  
  User::findBySql('SELECT * FROM user')->one();  //此方法是用 sql  语句查询 user 表里面的一条数据；
  
  User::find()->andWhere(['sex' => '男', 'age' => '24'])->count('id');   //统计符合条件的总条数；
  
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
//读取config/params.php的配置
 echo Yii::$app->params['name'];

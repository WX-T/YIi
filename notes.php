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
  

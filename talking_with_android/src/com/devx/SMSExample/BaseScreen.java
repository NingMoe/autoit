package com.devx.SMSExample;

import android.app.Activity;
import android.os.Bundle;
import android.view.KeyEvent;

public class BaseScreen extends Activity {
    /** Called when the activity is first created. */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.main);
        
    }
    public void onResume()
    {
    	super.onResume();
    	setContentView(R.layout.main);
    }
    public void onPause()
    {
    	super.onPause();
    }
    public void onDestroy()
    {
    	super.onDestroy();
    }
    
    
}
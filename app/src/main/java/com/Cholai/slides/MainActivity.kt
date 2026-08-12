package com.cholai.slides

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.Button
import android.widget.EditText
import androidx.appcompat.app.AppCompatActivity
import com.google.firebase.analytics.FirebaseAnalytics
import com.google.firebase.database.FirebaseDatabase

class MainActivity : AppCompatActivity() {
    
    private lateinit var firebaseAnalytics: FirebaseAnalytics
    private lateinit var database: FirebaseDatabase

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        
        firebaseAnalytics = FirebaseAnalytics.getInstance(this)
        database = FirebaseDatabase.getInstance()
        
        // Log app open for Cholai HQ
        val bundle = Bundle()
        bundle.putString("app_version", "1.1")
        firebaseAnalytics.logEvent("app_open", bundle)

        val btnWebsite: Button = findViewById(R.id.btnWebsite)
        val btnOrder: Button = findViewById(R.id.btnOrder)
        val editName: EditText = findViewById(R.id.editName)
        val editPhone: EditText = findViewById(R.id.editPhone)

        btnWebsite.setOnClickListener {
            val intent = Intent(Intent.ACTION_VIEW, Uri.parse("https://be708.github.io"))
            startActivity(intent)
        }

        btnOrder.setOnClickListener {
            val name = editName.text.toString()
            val phone = editPhone.text.toString()
            
            val orderData = mapOf("name" to name, "phone" to phone, "timestamp" to System.currentTimeMillis())
            database.reference.child("orders").push().setValue(orderData)
            
            val orderBundle = Bundle()
            orderBundle.putString("name", name)
            firebaseAnalytics.logEvent("order_submitted", orderBundle)
        }
    }
}

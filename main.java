
public class RegStatusActivity extends AppCompatActivity {
    private Spinner gm, rsm, dsm, spinner;

    private File directory, sd, file;
    private WritableWorkbook workbook;
    private List<all_data_model> list;
    private static final String TAG = "APP";
    private DatabaseReference databaseReference, datab;

    private String dsmname;
    private String smname;
    private String smffc;
    private String rsmname;
    private String rsmffc;
    private String rfid;
    private String region;

    private String userdsmcode;
    private String username;
    private String userphoneno;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_reg_status);

        databaseReference = FirebaseDatabase.getInstance().getReference().child("Users");
        datab = FirebaseDatabase.getInstance().getReference().child("data");
        ShowSpinner();
        list = new ArrayList<>();

        databaseReference.addValueEventListener(new ValueEventListener() {
            @Override
            public void onDataChange(@NonNull DataSnapshot dataSnapshot) {

                for (DataSnapshot ds : dataSnapshot.getChildren()) {

                    userdsmcode = String.valueOf(ds.child("userdsmcode").getValue());
                    username = String.valueOf(ds.child("username").getValue());
                    userphoneno = String.valueOf(ds.child("userphoneno").getValue());

                    AllThisInfo(userdsmcode, username, userphoneno);

                }


            }

            @Override
            public void onCancelled(@NonNull DatabaseError databaseError) {

            }
        });
    }

    private void AllThisInfo(final String userdsmcode, String username, String userphoneno) {

        final String dsm = userdsmcode;
        final String name = username;
        final String phone = userphoneno;
        datab.orderByChild("G").equalTo(dsm).addListenerForSingleValueEvent(new ValueEventListener() {
            @Override
            public void onDataChange(@NonNull DataSnapshot dataSnapshot) {


                for (DataSnapshot ds : dataSnapshot.getChildren()) {
                    dsmname = String.valueOf(ds.child("A").getValue());
                    smname = String.valueOf(ds.child("B").getValue());
                    smffc = String.valueOf(ds.child("C").getValue());
                    rsmname = String.valueOf(ds.child("D").getValue());
                    rsmffc = String.valueOf(ds.child("E").getValue());
                    rfid = String.valueOf(ds.child("F").getValue());
                    region = String.valueOf(ds.child("H").getValue());
                }

                list.add(new all_data_model(dsm, name, phone, dsmname, smname, smffc, rsmname, rsmffc, rfid, userdsmcode, region));


            }

            @Override
            public void onCancelled(@NonNull DatabaseError databaseError) {

            }
        });


    }



    public void DownLoadFile(View view) {          //I just added a button in my xml file 

        createExcelSheet();
    }

    public void createExcelSheet() {
        if (isStoragePermissionGranted()) {
            String csvFile = "FullUserInfo.xls";
            sd = Environment.getExternalStorageDirectory();
            directory = new File(sd.getAbsolutePath());
            file = new File(directory, csvFile);
            WorkbookSettings wbSettings = new WorkbookSettings();
            wbSettings.setLocale(new Locale("en", "EN"));
            try {
                workbook = Workbook.createWorkbook(file, wbSettings);
                createFirstSheet();

                //closing cursor
                workbook.write();
                workbook.close();

                Toast.makeText(this, "File Downloaded !", Toast.LENGTH_SHORT).show();
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {

            Toast.makeText(this, "Permission Denied !", Toast.LENGTH_SHORT).show();
        }
    }

    public void createFirstSheet() {
        try {

            //Excel sheet name. 0 (number)represents first sheet
            WritableSheet sheet = workbook.createSheet("sheet1", 0);
            // column and row title
            sheet.addCell(new Label(0, 0, "DSMCODE"));
            sheet.addCell(new Label(1, 0, "USERNAME"));
            sheet.addCell(new Label(2, 0, "PHONE_NO"));
            sheet.addCell(new Label(3, 0, "DSM_NAME"));
            sheet.addCell(new Label(4, 0, "SM_NAME"));
            sheet.addCell(new Label(5, 0, "SM_FFC"));
            sheet.addCell(new Label(6, 0, "RSMNAME"));
            sheet.addCell(new Label(7, 0, "RSM_FFC"));
            sheet.addCell(new Label(8, 0, "RFID"));
            sheet.addCell(new Label(9, 0, "FFC"));
            sheet.addCell(new Label(10, 0, "REGION"));


            for (int i = 0; i < list.size(); i++) {
                sheet.addCell(new Label(0, i + 1, list.get(i).getUserdsmcode()));
                sheet.addCell(new Label(1, i + 1, list.get(i).getUsername()));
                sheet.addCell(new Label(2, i + 1, list.get(i).getUserphoneno()));
                sheet.addCell(new Label(3, i + 1, list.get(i).getA()));
                sheet.addCell(new Label(4, i + 1, list.get(i).getB()));
                sheet.addCell(new Label(5, i + 1, list.get(i).getC()));
                sheet.addCell(new Label(6, i + 1, list.get(i).getD()));
                sheet.addCell(new Label(7, i + 1, list.get(i).getE()));
                sheet.addCell(new Label(8, i + 1, list.get(i).getF()));
                sheet.addCell(new Label(9, i + 1, list.get(i).getG()));
                sheet.addCell(new Label(10, i + 1, list.get(i).getH()));

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public boolean isStoragePermissionGranted() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            if (checkSelfPermission(android.Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    == PackageManager.PERMISSION_GRANTED) {
                Log.v(TAG, "Permission is granted");
                return true;
            } else {

                Log.v(TAG, "Permission is revoked");
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
                return false;
            }
        } else { //permission is automatically granted on sdk<23 upon installation
            Log.v(TAG, "Permission is granted");
            return true;
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String[] permissions, int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            Log.v(TAG, "Permission: " + permissions[0] + "was " + grantResults[0]);
            //resume tasks needing this permission
        }
    }
}

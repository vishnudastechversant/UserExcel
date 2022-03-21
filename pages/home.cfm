<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Manager</title>
    <link rel="stylesheet" href="../assets/css/style.css">
    <link rel="stylesheet" href="../assets/bootstrap5/css/bootstrap.css">
</head>
<body>
    <section>
        <div class="d-flex justify-content-around mt-5 p-3">
            <div class="col d-flex justify-content-center"><button class="btn btn-secondary" onclick="check('Plane Template')">Plane Template</button></div>
            <div class="col d-flex justify-content-center"><button class="btn btn-primary" onclick="check('Template with data')">Template With Data</button></div>
            <div class="col d-flex justify-content-center"><button class="btn btn-dark" onclick="check('Browse')">Browse</button></div>
            <div class="col d-flex justify-content-center"><button class="btn btn-success" onclick="check('Upload')">Upload</button></div>
            <div class="col d-flex justify-content-center"><button class="btn btn-danger" onclick="check('Download')">Download</button></div>
        </div>
        <div class="d-flex justify-content-center">
            <table class="table table-bordered bg-light m-5">
                <thead>
                    <tr>
                        <th>First Name</th>
                        <th>Last Name</th>
                        <th>Address</th>
                        <th>Email</th>
                        <th>Phone</th>
                        <th>DOB</th>
                        <th>Role</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>fname</td>
                        <td>fname</td>
                        <td>fname</td>
                        <td>fname</td>
                        <td>fname</td>
                        <td>fname</td>
                        <td>fname</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </section>
    <script>
        const check = (text) =>{
            alert(text);
        }
    </script>
    <script src="../assets/bootstrap5/js/bootstrap.js"></script>
</body>
</html>
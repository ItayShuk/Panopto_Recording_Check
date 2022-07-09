<!-- ABOUT THE PROJECT -->
## About The Project
This program checking validate Panopto servers status are recording.
Doing so by extracting the lectures data from .xml file in its directory,
checking which hall should be recording,
then validating it's recording status.
If any of the servers isn't recording, we will get notify in mail.

<!-- GETTING STARTED -->
## Getting Started


### Prerequisites

Use Python 3.8, and run
```sh
pip install -r requirements.txt

```

### Installation

1. Sign in to the Panopto web site as Administarator
2. Click the System icon at the left-bottom corner.
3. Click API Clients
4. Click New
5. Enter arbitrary Client Name
6. Select Server-side Web Application type.
7. Enter https://localhost into CORS Origin URL.
8. Enter http://localhost:9127/redirect into Redirect URL.
9. The rest can be blank. Click "Create API Client" button.
10. Note the created Client ID and Client Secret.



<!-- USAGE EXAMPLES -->
## Usage
The client id and client secret are necessary. If you provide only them, all the database will be migrated.
You can add course id, year, semester. In this case only what you entered will be migrated.
In addition, you can add folder id. In this case, what you ask to upload will be uploaded to this specific panopto folder id.

```

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.



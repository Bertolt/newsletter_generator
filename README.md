# Newsletter Generator
This projects aims to create an automated tool to create a newsletter in HTML format to embedded on Outlook email.


## Getting Started

add git push commands for repository

### Prerequisites

Python 3.7.0 interpreter with the following packages (add: pip commands):
- Pandas
- Numpy
- Os
- Time
- Shutil
- Xlrd


Optional: 
- Docker: Image available
- Virtualenv: Pipenv

### Installing

Python 3.7.0:

1. Install Python interpreter
2. Create Working Directory
3. Push repository to working Directory: HTML and CSS templates, Python script, Excel sheet.

Docker:
1. Install Docker
2. Select Linux Containers
3. Create shared driver and and working directory for mount purposes
```
$ mkdir path/to/dir/mount_dir
```
4. Load image from .tar file: 
```
$ docker image load -i <path/to/image>/newsletter.tar
```
5. Push repository to mounted dir: (add git command)
6. Run Container:
```
$ docker run -it -v d:/m2:/app newsletter
```
or
```
$ docker run -it -v d:/m2:/app newsletter /bin/bash
``` 
7. Container mounted directory: 
```
$ /app
```

## Running the tests

Console:
```
$ python path/to/dir/news_generator.py
```

Docker:
```
$ docker run -it -v d:/m2:/app newsletter
```
or
```
$ docker run -it -v d:/m2:/app newsletter /bin/bash
``` 

### Break down into end to end tests

The script reads the templates and appends the information according to the specified in the excel file.
Finally it generates the  newsletter.html


### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Refer to installing notes

## Built With

* [Python](http://www.dropwizard.io/1.0.2/docs/) - Language and interpreter
* [Pycharm](https://maven.apache.org/) - IDE and debugger
* [Docker](https://rometools.github.io/rome/) - Used to create container for multiple platforms and security

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

<<<<<<< HEAD
=======
## Versioning

We use Git(http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 
>>>>>>> 870cd18... first commit

## Authors

* **Ruben Bertelo** - *Initial work* - [Rbertolt](https://github.com/PurpleBooth)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

<<<<<<< HEAD
=======
## Acknowledgments

* Hat tip to anyone whose code was used
* Inspiration
* etc
>>>>>>> 870cd18... first commit


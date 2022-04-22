## Release

* Change version in `build.gradle` file
* Run `gradle smokeTest` to run all the tests
* Run `gradle distZip`. Zip file will be in `build/distributions` folder
* Upload release to github
* Update brew formula https://docs.brew.sh/How-To-Open-a-Homebrew-Pull-Request#submit-a-new-version-of-an-existing-formula: `brew bump-formula-pr excel-compare --url <new_release.zip> --sha256 <sha256 of new_release.zip>`

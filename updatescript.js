const gitHubPath = 'agam778/Microsoft-Office-Electron';
const url = 'https://api.github.com/repos/' + gitHubPath + '/tags';

$.get(url).done(data => {
  const versions = data.sort((v1, v2) => semver.compare(v2.name, v1.name));
  $('#result').html(versions[0].name);
});

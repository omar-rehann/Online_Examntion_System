<?php

use Omar\Omar; 
use Omar\Version\GitVersionCollection;
use Symfony\Component\Finder\Finder;

$iterator = Finder::create()
    ->files()
    ->name('*.php')
    ->in($dir = __DIR__ . '/src');

$versions = GitVersionCollection::create($dir)
    ->addFromTags(function ($version) {
        return preg_match('~^\d+\.\d+\.\d+$~', $version);
    })
    ->add('master');

return new Omar($iterator, [
    'title' => 'omar_rehan/online_exam_system', 
    'versions' => $versions,
    'build_dir' => __DIR__ . '/build/%version%',
    'cache_dir' => __DIR__ . '/cache/%version%',
    'remote_repository' => new GitHubRemoteRepository('YourGitHubUsername/YourRepoName', dirname($dir)),
]);

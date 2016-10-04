var app = angular.module('appMovies',[])
var Controllers = {};
    Controllers.newMovie = function ($scope,$http) {
    $scope.save = function () {
        console.log($scope.Movie);
        $scope.result = JSON.stringify($scope.Movie);
    }
}
app.controller(Controllers);
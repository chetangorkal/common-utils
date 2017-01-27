var jnvcAlumniApp = angular.module('jnvcAlumniApp',[ 'ngRoute' ]);

jnvcAlumniApp.config([ '$routeProvider', function($routeProvider) {
	$routeProvider.when('/', {
		controller : 'HomeController'
	}).when('/about', {
		templateUrl : 'resources/view/booking.html',
		controller : 'HomeController'
	}).when('/Contact', {
		controller : 'HomeController'
	}).otherwise({
		redirectTo : '/'
	});
} ]);

jnvcAlumniApp.controller('HomeController', function($scope, $http){
	
	$http({
        method : "GET",
        url : "jnvc_alumni_data.json"
    }).then(function mySucces(response) {
        $scope.data = response.data;
    }, function myError(response) {
        $scope.data = response.statusText;
    });
	
});

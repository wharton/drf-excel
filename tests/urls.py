from rest_framework import routers

from .testapp.views import AllFieldsViewSet, ExampleViewSet, SecretFieldViewSet

router = routers.SimpleRouter()
router.register(r"examples", ExampleViewSet)
router.register(r"all-fields", AllFieldsViewSet)
router.register(r"secret-field", SecretFieldViewSet)

urlpatterns = router.urls

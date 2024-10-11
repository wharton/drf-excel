from rest_framework import routers

from .testapp.views import ExampleViewSet, AllFieldsViewSet

router = routers.SimpleRouter()
router.register(r"examples", ExampleViewSet)
router.register(r"all-fields", AllFieldsViewSet)

urlpatterns = router.urls

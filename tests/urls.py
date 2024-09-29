from rest_framework import routers

from .testapp.views import ExampleViewSet

router = routers.SimpleRouter()
router.register(r'examples', ExampleViewSet)

urlpatterns = router.urls

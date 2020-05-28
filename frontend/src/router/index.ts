import Vue from "vue";
import VueRouter, { RouteConfig } from "vue-router";
import Transform from "../views/Transform.vue";

Vue.use(VueRouter);

const routes: Array<RouteConfig> = [
  {
    path: "/transform",
    name: "transform",
    component: Transform
  },
  {
    path: "/check",
    name: "Check",
    // route level code-splitting
    // this generates a separate chunk (about.[hash].js) for this route
    // which is lazy-loaded when the route is visited.
    component: () =>
      import(/* webpackChunkName: "about" */ "../views/Check.vue")
  }
];

const router = new VueRouter({
  mode: "hash",
  base: process.env.BASE_URL,
  routes
});

export default router;

import Vue from "vue";
import VueRouter, { RouteConfig } from "vue-router";
import Split from "../views/Split.vue";

Vue.use(VueRouter);

const routes: Array<RouteConfig> = [
  {
    path: "/",
    redirect: "/split"
  },
  {
    path: "/split",
    name: "split",
    component: Split
  },
  {
    path: "/evaluation",
    name: "Evaluation",
    // route level code-splitting
    // this generates a separate chunk (about.[hash].js) for this route
    // which is lazy-loaded when the route is visited.
    component: () =>
      import(/* webpackChunkName: "about" */ "../views/Evaluation.vue")
  }
];

const router = new VueRouter({
  mode: "hash",
  base: process.env.BASE_URL,
  routes
});

export default router;
